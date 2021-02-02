# Google Sheets feedback

The Google Sheets feedback solution allows a team to ask all team members to mutually rate each other on pre-defined criteria for feedback talks. The result Google Sheet provides the team's average rating, standard deviation and comments for every team member.

## Procedure

The Python `Feedback` class has two main methods: `create_feedback_sheets()` creates new Google Sheets for all team members to enter mutual feedback; `evaluate_feedback_sheets()` collects the feedback and distributes the results into a feedback Google Sheet for every team member.

## Create service account

Create or choose a project in the Google [cloud console](https://console.cloud.google.com/). Enable the [Google Sheet API](https://console.cloud.google.com/apis/api/sheets.googleapis.com/overview). Create a [service account](https://console.cloud.google.com/projectselector2/iam-admin/serviceaccounts?supportedpurview=project) and grant permissions on the Google Sheet API.

Update `app/service-account.json` with the service account credentials.

## Create feedback template

Make a copy of the [Feedback Template Example](https://docs.google.com/spreadsheets/d/1pJTw1gYVLp1j-BCfz1qXD3IgMX2UwYO90nTeFZ8FaOE/edit#gid=519050822) sheet. Share the copy with the service account. Adjust categories and topics in the template `Input` and `Results` tabs. Enter names of team members in the `Names` tab.

The `master_users` entry in `app/feedback.json` contains the accounts which will have access to all generated Google Sheets to supervise the process. If you change the layout of the template sheets you need to update the `input` and `results` entries.