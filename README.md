# MB Inventory System

Detailed documentation for this project can be found in [PROJECT_DOCUMENTATION.md](PROJECT_DOCUMENTATION.md).

## Setup Instructions

Setting up the project locally:

1. Clone this repository.

2. Run `clasp login` to authenticate your Google account.

3. Do one of the following:

 - **To create a new script:** Run `clasp create --type standalone`. This will automatically generate a valid `.clasp.json` file for you.

 - **To link an existing script:** Copy `.clasp.json.template` to `.clasp.json` and replace "SCRIPT_ID_HERE" with the actual Google Apps Script ID.

4. Run `clasp push` to deploy the code.
