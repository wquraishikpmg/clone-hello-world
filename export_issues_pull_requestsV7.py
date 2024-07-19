import requests
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Color
from openpyxl.styles.colors import BLUE

#repo_owner = "wquraishikpmg"
#repo_name = "clone-hello-world"
# to be tested with below as has lots of data
#repo_owner = "microsoft"
#repo_name = "azurechat"


#repo_owner = "wquraishikpmg"
#repo_name = "hello-world"
access_token = "" 

# Fetch issues and pull requests using GitHub REST API
issues_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues"
pulls_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/pulls"

headers = {"Authorization": f"Bearer {access_token}"}

# Fetch issues
issues_response = requests.get(issues_url, headers=headers)
if issues_response.status_code == 200:
    issues_data = issues_response.json()
else:
    print(f"Failed to fetch issues. Error: {issues_response.text}")
    issues_data = []

# Fetch pull requests
pulls_response = requests.get(pulls_url, headers=headers)
if pulls_response.status_code == 200:
    pulls_data = pulls_response.json()
else:
    print(f"Failed to fetch pull requests. Error: {pulls_response.text}")
    pulls_data = []

# Create an Excel workbook
wb = Workbook()
sheet = wb.active
sheet.title = "Issues and Pull Requests"

# Set header row font to bold
for cell in sheet[1]:
    cell.font = Font(bold=True)

# Write the header row
header_row = ["Number", "Type", "Title", "Body", "Assignees", "Labels", "Milestone", "State", "Reviewers", "Committers"]
for col_num, header in enumerate(header_row, 1):
    col_letter = get_column_letter(col_num)
    sheet[f"{col_letter}1"] = header

# 
# Write issues data to the Excel file
for row_num, issue in enumerate(issues_data, 2):
    issue_number = issue["number"]
    issue_type = "Issue"
    issue_title = issue["title"]
    issue_body = issue["body"]
    issue_assignees = ",".join(assignee["login"] for assignee in issue["assignees"])
    issue_labels = ",".join(label["name"] for label in issue["labels"])
    issue_milestone = issue["milestone"]["title"] if issue["milestone"] else ""
    issue_state = issue["state"]
    issue_url = issue["html_url"]

    cell = f"A{row_num}"
    sheet[cell].hyperlink = f'{issue_url}'
    sheet[cell].value = issue_number
    sheet[cell].font = Font(color=Color(rgb=BLUE))
    sheet[f"B{row_num}"] = issue_type
    sheet[f"C{row_num}"] = issue_title
    sheet[f"D{row_num}"] = issue_body
    sheet[f"E{row_num}"] = issue_assignees
    sheet[f"F{row_num}"] = issue_labels
    sheet[f"G{row_num}"] = issue_milestone
    sheet[f"H{row_num}"] = issue_state

# Write pull requests data to the Excel file
for pull_num, pull in enumerate(pulls_data, row_num + 1):
    pull_number = pull["number"]
    pull_type = "Pull Request"
    pull_title = pull["title"]
    pull_body = pull["body"]
    pull_assignees = ",".join(assignee["login"] for assignee in pull["assignees"])
    pull_labels = ",".join(label["name"] for label in pull["labels"])
    pull_milestone = pull["milestone"]["title"] if pull["milestone"] else ""
    pull_state = pull["state"]
    pull_url = pull["html_url"]
    pull_reviewers = ",".join(reviewer["login"] for reviewer in pull["requested_reviewers"])
    pull_committers = pull["user"]["login"]

    cell = f"A{pull_num}"
    sheet[cell].hyperlink = f'{pull_url}'
    sheet[cell].value = pull_number
    sheet[cell].font = Font(color=Color(rgb=BLUE))
    sheet[f"B{pull_num}"] = pull_type
    sheet[f"C{pull_num}"] = pull_title
    sheet[f"D{pull_num}"] = pull_body
    sheet[f"E{pull_num}"] = pull_assignees
    sheet[f"F{pull_num}"] = pull_labels
    sheet[f"G{pull_num}"] = pull_milestone
    sheet[f"H{pull_num}"] = pull_state
    sheet[f"I{pull_num}"] = pull_reviewers
    sheet[f"J{pull_num}"] = pull_committers

# Adjust column widths
for col in sheet.columns:
    max_length = 0
    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    adjusted_width = (max_length + 2) * 1.2
    sheet.column_dimensions[col[0].column_letter].width = adjusted_width

# Save the workbook as an Excel file
wb.save("issues_and_pull_requests.xlsx")
