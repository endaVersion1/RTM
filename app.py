import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import datetime

# RTM columns
RTM_COLUMNS = [
    "Jira ID", "Jira Summary", "Test Case IDs", "Test Case Titles", "Test Case Status",
    "Sprint", "Jira Status", "Jira Updated", "Description", "Acceptance Criteria",
    "Jira Link", "Test Case Links", "Tested On"
]

JIRA_URL = "https://digitalhubjira.atlassian.net/issues/{}"
TESTRAIL_URL = "https://dotdigitalhub.testrail.io/index.php?/tests/view/{}"

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
        tc_ids = ", ".join(tc_matches["ID"].astype(str)) if not tc_matches.empty else ""
        tc_titles = ", ".join(tc_matches["Title"].astype(str)) if not tc_matches.empty else ""
        tc_status = ", ".join(tc_matches["Status"].astype(str)) if not tc_matches.empty else ""
        tc_tested_on = ", ".join(tc_matches["Tested On"].astype(str)) if not tc_matches.empty else ""
        # Links
        jira_link = JIRA_URL.format(jira_id) if jira_id else ""
        tc_links = ", ".join([TESTRAIL_URL.format(tc_id.lstrip("T")) for tc_id in tc_matches["ID"]]) if not tc_matches.empty else ""
        rtm_rows.append([
            jira_id, summary, tc_ids, tc_titles, tc_status, sprint, status, updated, desc, ac, jira_link, tc_links, tc_tested_on
        ])
    rtm_df = pd.DataFrame(rtm_rows, columns=RTM_COLUMNS)
    return rtm_df

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
            if row[10]:  # Jira Link
                worksheet.write_url(row_num, 10, row[10], string=row[0])
            if row[11]:  # Test Case Links
                # If multiple links, only link the first
                first_link = row[11].split(", ")[0]
                first_tc_id = row[2].split(", ")[0] if row[2] else ""
                worksheet.write_url(row_num, 11, first_link, string=first_tc_id)
    output.seek(0)
    return output

def main():
    st.title("Requirement Traceability Matrix Generator")
    st.write("Upload Jira and TestRail CSV files to generate an RTM Excel document.")
    jira_file = st.file_uploader("Upload Jira CSV", type=["csv"])
    testrail_file = st.file_uploader("Upload TestRail CSV", type=["csv"])
    if jira_file and testrail_file:
        jira_df = load_csv(jira_file)
        testrail_df = load_csv(testrail_file)
        rtm_df = generate_rtm(jira_df, testrail_df)
        st.dataframe(rtm_df)
        excel_data = to_excel(rtm_df)
        st.download_button(
            label="Download RTM Excel",
            data=excel_data,
            file_name=f"RTM_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
