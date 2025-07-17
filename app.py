import datetime
from io import BytesIO

import pandas as pd
import requests
from datetime import datetime
import streamlit as st

# RTM columns
RTM_COLUMNS = [
    "Jira ID", "Jira Summary", "Test Case IDs", "Test Case Titles", "Test Case Status",
    "Sprint", "Jira Status", "Jira Updated", "Description", "Acceptance Criteria",
    "Jira Link", "Test Case Links", "Tested On"
]

JIRA_URL = "https://digitalhubjira.atlassian.net/issues/{}"
TESTRAIL_URL = "https://dotdigitalhub.testrail.io/index.php?/api/v2"

STATUS_MAP = {
    1: "Passed",
    2: "Blocked",
    3: "Untested",
    4: "Retest",
    5: "Failed"
}

# --- API Integration Functions ---
def fetch_jira_issues(jira_url, jira_email, jira_token, jira_jql):
    headers = {
        "Accept": "application/json"
    }
    auth = (jira_email, jira_token)
    params = {
        "jql": jira_jql,
        "maxResults": 100
    }
    response = requests.get(f"{jira_url}/rest/api/2/search", headers=headers, auth=auth, params=params)
    response.raise_for_status()
    data = response.json()
    issues = data["issues"]
    # Flatten issues for DataFrame
    flattened_issues = []
    for issue in issues:
        fields = issue["fields"]
        # Convert updated_on (epoch) to date string if present
        updated_on = fields.get("updated_on", "")
        if updated_on:
            try:
                updated_on = datetime.fromtimestamp(updated_on).strftime("%Y-%m-%d")
            except Exception:
                pass
        flattened_issues.append({
            "Issue key": issue["key"],
            "Summary": fields.get("summary", ""),
            "Status": fields.get("status", {}).get("name", ""),
            "Updated": updated_on,
            "Sprint": next((s.get("name", "") for s in fields.get("customfield_10020", []) if isinstance(s, dict)), ""),
            "Description": fields.get("description", ""),
            "Custom field (Acceptance Criteria)": fields.get("customfield_10021", "")
        })
    return pd.DataFrame(flattened_issues)

def fetch_testrail_cases(testrail_url, username, api_key):
    headers = {"Content-Type": "application/json"}
    response = requests.get(
        f"{testrail_url}/index.php?/api/v2/get_cases/2",
        auth=(username, api_key), headers=headers
    )
    response.raise_for_status()
    return response.json()

def load_csv(file):
    return pd.read_csv(file)

def generate_rtm(jira_df, testrail_df):
    # Map TestRail test cases to Jira issues by searching for Jira ID in Title
    rtm_rows = []
    for _, jira_row in jira_df.iterrows():
        jira_id = jira_row.get("Issue key") or jira_row.get("Jira ID")
        summary = jira_row.get("Summary")
        status = jira_row.get("Status")
        updated = jira_row.get("Updated")
        sprint = jira_row.get("Sprint")
        desc = jira_row.get("Description")
        ac = jira_row.get("Custom field (Acceptance Criteria)")
        # Find matching test cases by Jira ID in Title
        tc_matches = testrail_df[testrail_df["Title"].str.contains(str(jira_id), na=False)]
        # Only join non-empty values
        tc_ids = ", ".join([v for v in tc_matches["ID"].astype(str) if v.strip()]) if not tc_matches.empty else ""
        tc_titles = ", ".join([v for v in tc_matches["Title"].astype(str) if v.strip()]) if not tc_matches.empty else ""
        tc_status = ", ".join([v for v in tc_matches["Status"].astype(str) if v.strip()]) if not tc_matches.empty else ""
        tc_tested_on = ", ".join([v for v in tc_matches["Tested On"].astype(str) if v.strip()]) if not tc_matches.empty else ""
        # Links
        jira_link = JIRA_URL.format(jira_id) if jira_id else ""
        tc_links = []
        for tc_id in tc_matches["ID"]:
            # Remove leading 'T' if present
            clean_id = str(tc_id).lstrip("T")
            tc_links.append(f"https://dotdigitalhub.testrail.io/index.php?/tests/view/{clean_id}")
        tc_links_str = ", ".join(tc_links) if tc_links else ""
        rtm_rows.append({
            "Jira ID": jira_id,
            "Jira Summary": summary,
            "Test Case IDs": tc_ids,
            "Test Case Titles": tc_titles,
            "Test Case Status": tc_status,
            "Tested On": tc_tested_on,
            "Sprint": sprint,
            "Jira Status": status,
            "Jira Updated": updated,
            "Description": desc,
            "Acceptance Criteria": ac,
            "Jira Link": jira_link,
            "Test Case Links": tc_links_str
        })
    return pd.DataFrame(rtm_rows)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="RTM")
        workbook = writer.book
        worksheet = writer.sheets["RTM"]
        # Add autofilter
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        # Add hyperlinks
        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            # Jira Link
            if row[10] and str(row[10]).startswith(('http://', 'https://')):
                worksheet.write_url(row_num, 10, row[10], string=row[0])
            else:
                worksheet.write(row_num, 10, row[10])
            # Test Case Links
            if row[11]:
                first_link = row[11].split(", ")[0]
                first_tc_id = row[2].split(", ")[0] if row[2] else ""
                if first_link.startswith(('http://', 'https://')):
                    worksheet.write_url(row_num, 11, first_link, string=first_tc_id)
                else:
                    worksheet.write(row_num, 11, first_link)
    output.seek(0)
    return output

def to_html(df):
    df_html = df.copy()
    # Make Jira Link clickable
    if "Jira Link" in df_html.columns and "Jira ID" in df_html.columns:
        df_html["Jira Link"] = df_html.apply(
            lambda x: f'<a href="{x["Jira Link"]}" target="_blank">{x["Jira ID"]}</a>' if x["Jira Link"] else "",
            axis=1
        )
    # Make Test Case Links clickable
    if "Test Case Links" in df_html.columns and "Test Case IDs" in df_html.columns:
        def make_tc_links(row):
            links = str(row["Test Case Links"]).split(", ")
            ids = str(row["Test Case IDs"]).split(", ")
            return ", ".join([
                f'<a href="{link}" target="_blank">{tc_id}</a>' if link and tc_id else "" for link, tc_id in zip(links, ids)
            ]) if links and ids else ""
        df_html["Test Case Links"] = df_html.apply(make_tc_links, axis=1)
    html = df_html.to_html(index=False, escape=False, justify="center", border=1)
    return html.encode("utf-8")

def fetch_testrail_cases_and_statuses(testrail_url, username, api_key, project_id, suite_id=None):
    headers = {"Content-Type": "application/json"}
    # 1. Get all plans for the project
    plans_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_plans/{project_id}", auth=(username, api_key), headers=headers)
    plans_resp.raise_for_status()
    plans_data = plans_resp.json()
    plans = plans_data['plans'] if isinstance(plans_data, dict) and 'plans' in plans_data else plans_data
    case_status_map = {}
    case_title_map = {}
    case_updatedon_map = {}
    # 2. Loop through each plan and its runs
    for plan in plans:
        plan_id = plan.get("id")
        plan_detail_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_plan/{plan_id}", auth=(username, api_key), headers=headers)
        if plan_detail_resp.status_code != 200:
            continue
        plan_detail = plan_detail_resp.json()
        for entry in plan_detail.get("entries", []):
            for run in entry.get("runs", []):
                run_id = run.get("id")
                tests_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_tests/{run_id}", auth=(username, api_key), headers=headers)
                if tests_resp.status_code != 200:
                    continue
                tests = tests_resp.json()
                if isinstance(tests, dict) and 'tests' in tests:
                    tests = tests['tests']
                for test in tests:
                    case_id = test.get("case_id")
                    status_id = test.get("status_id", "")
                    updated_on = test.get("updated_on", 0)
                    # Map status_id to text
                    status_text = STATUS_MAP.get(status_id, str(status_id)) if status_id else ""
                    # Get case title (only once per case_id)
                    if case_id and case_id not in case_title_map:
                        case_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_case/{case_id}", auth=(username, api_key), headers=headers)
                        if case_resp.status_code == 200:
                            case_data = case_resp.json()
                            case_title_map[case_id] = case_data.get("title", "")
                    # Only keep the latest status for each case_id
                    if case_id and (case_id not in case_status_map or updated_on > case_status_map[case_id][1]):
                        case_status_map[case_id] = (status_text, updated_on)
                        case_updatedon_map[case_id] = updated_on
    # Build enriched cases list
    enriched_cases = []
    for case_id, (status, updated_on) in case_status_map.items():
        title = case_title_map.get(case_id, "")
        tested_on = ""
        if updated_on:
            tested_on = datetime.datetime.fromtimestamp(updated_on).strftime("%Y-%m-%d")
        enriched_cases.append({
            "ID": str(case_id),
            "Title": title,
            "Status": status,
            "Tested On": tested_on
        })
    return enriched_cases

def fetch_testrail_cases_status_testedon(testrail_url, username, api_key, project_id):
    headers = {"Content-Type": "application/json"}
    # 1. Get all plans for the project
    plans_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_plans/2", auth=(username, api_key), headers=headers)
    plans_resp.raise_for_status()
    plans = plans_resp.json()
    if isinstance(plans, dict) and 'plans' in plans:
        plans = plans['plans']
    case_status_map = {}
    case_testedon_map = {}
    case_title_map = {}
    # 2. Loop through each plan and its runs
    for plan in plans:
        plan_id = plan.get("id")
        # Get full plan details (to get runs)
        plan_detail_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_plan/{plan_id}", auth=(username, api_key), headers=headers)
        if plan_detail_resp.status_code != 200:
            continue
        plan_detail = plan_detail_resp.json()
        for entry in plan_detail.get("entries", []):
            for run in entry.get("runs", []):
                run_id = run.get("id")
                # Get all tests in this run
                tests_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_tests/{run_id}", auth=(username, api_key), headers=headers)
                if tests_resp.status_code != 200:
                    continue
                tests = tests_resp.json()
                if isinstance(tests, dict) and 'tests' in tests:
                    tests = tests['tests']
                for test in tests:
                    case_id = test.get("case_id")
                    status_id = test.get("status_id", "")
                    # Map status_id to text
                    status_text = STATUS_MAP.get(status_id, str(status_id)) if status_id else ""
                    # Get case details for updated_on and title
                    if case_id not in case_title_map:
                        case_resp = requests.get(f"{testrail_url}/index.php?/api/v2/get_case/{case_id}", auth=(username, api_key), headers=headers)
                        if case_resp.status_code == 200:
                            case_data = case_resp.json()
                            case_title_map[case_id] = case_data.get("title", "")
                            updated_on = case_data.get("updated_on", "")
                            if updated_on:
                                updated_on = datetime.datetime.fromtimestamp(updated_on).strftime("%Y-%m-%d")
                                case_testedon_map[case_id] = updated_on
                    # Always use the latest status_id for the case (by test's updated_on)
                    if case_id not in case_status_map or test.get("updated_on", 0) > case_status_map[case_id][1]:
                        case_status_map[case_id] = (status_text, test.get("updated_on", 0))
    # Build enriched cases list
    enriched_cases = []
    for case_id, (status, _) in case_status_map.items():
        enriched_cases.append({
            "ID": str(case_id),
            "Title": case_title_map.get(case_id, ""),
            "Status": status,
            "Tested On": case_testedon_map.get(case_id, "")
        })
    return enriched_cases

def get_enriched_test_case_data(project_id=2):
    TESTRAIL_BASE = "https://dotdigitalhub.testrail.io/index.php?/api/v2"
    AUTH = ("username", "api_key")  # Replace with actual credentials
    headers = {"Content-Type": "application/json"}
    enriched_cases = []

    # 1. Get all test plans
    plans = requests.get(f"{TESTRAIL_BASE}/get_plans/{project_id}", auth=AUTH, headers=headers).json()

    for plan in plans:
        plan_id = plan["id"]

        # 2. Get full plan details (to access runs)
        plan_detail = requests.get(f"{TESTRAIL_BASE}/get_plan/{plan_id}", auth=AUTH, headers=headers).json()

        for entry in plan_detail.get("entries", []):
            for run in entry.get("runs", []):
                run_id = run["id"]

                # 3. Get all tests from the run
                tests = requests.get(f"{TESTRAIL_BASE}/get_tests/{run_id}", auth=AUTH, headers=headers).json()
                for test in tests.get("tests", tests):  # handle both list and dict format
                    case_id = test["case_id"]
                    status_id = test.get("status_id")
                    status = STATUS_MAP.get(status_id, "Unknown")

                    # 4. Get test case details
                    case_data = requests.get(f"{TESTRAIL_BASE}/get_case/{case_id}", auth=AUTH, headers=headers).json()
                    updated_on = datetime.fromtimestamp(case_data["updated_on"]).strftime("%Y-%m-%d")
                    title = case_data["title"]

                    enriched_cases.append({
                        "Case ID": case_id,
                        "Title": title,
                        "Status": status,
                        "Updated On": updated_on,
                        "Test Run ID": run_id
                    })
    return enriched_cases

def main():
    st.title("Requirement Traceability Matrix Generator")
    st.write("Upload Jira and TestRail CSV files or fetch from API to generate an RTM Excel and HTML document.")
    source = st.radio("Select data source", ["CSV Upload", "API Integration"])
    if source == "CSV Upload":
        jira_file = st.file_uploader("Upload Jira CSV", type=["csv"])
        testrail_file = st.file_uploader("Upload TestRail CSV", type=["csv"])
        if jira_file and testrail_file:
            jira_df = load_csv(jira_file)
            testrail_df = load_csv(testrail_file)
            rtm_df = generate_rtm(jira_df, testrail_df)
            st.dataframe(rtm_df)
            excel_data = to_excel(rtm_df)
            html_data = to_html(rtm_df)
            st.download_button(
                label="Download RTM Excel",
                data=excel_data,
                file_name=f"RTM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label="Download RTM HTML",
                data=html_data,
                file_name=f"RTM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                mime="text/html"
            )
    else:
        st.subheader("Jira API Settings")
        jira_url = st.text_input("Jira Base URL", "https://digitalhubjira.atlassian.net")
        jira_email = st.text_input("Jira Email")
        jira_token = st.text_input("Jira API Token", type="password")
        jira_jql = st.text_input("Jira JQL Query", "project = MCR AND type = Story AND status = 'QA Finished' ORDER BY created DESC")
        st.subheader("TestRail API Settings")
        testrail_url = st.text_input("TestRail Base URL", "https://dotdigitalhub.testrail.io")
        testrail_user = st.text_input("TestRail Email/Username")
        testrail_key = st.text_input("TestRail API Token", type="password")
        if st.button("Fetch and Generate RTM"):
            with st.spinner("Fetching data from APIs..."):
                try:
                    jira_df = fetch_jira_issues(jira_url, jira_email, jira_token, jira_jql)
                except requests.exceptions.HTTPError as e:
                    st.error(f"Jira API Error: {e}\nJQL used: {jira_jql}")
                    return
                enriched_cases = fetch_testrail_cases_and_statuses(testrail_url, testrail_user, testrail_key, 2, 2)
                testrail_df = pd.DataFrame(enriched_cases)
                st.write('TestRail Data Preview:', testrail_df.head(10))
                rtm_df = generate_rtm(jira_df, testrail_df)
                st.dataframe(rtm_df)
                excel_data = to_excel(rtm_df)
                html_data = to_html(rtm_df)
                st.download_button(
                    label="Download RTM Excel",
                    data=excel_data,
                    file_name=f"RTM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.download_button(
                    label="Download RTM HTML",
                    data=html_data,
                    file_name=f"RTM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                    mime="text/html"
                )

if __name__ == "__main__":
    main()
