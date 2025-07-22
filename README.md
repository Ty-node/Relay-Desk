# IT Ticketing System: Slack & Google Sheets Integration 🤖

---

## 📋 Table of Contents

- [📖 Introduction](#-introduction)
- [✅ Prerequisites](#-prerequisites)
- [🚀 Step-by-Step Implementation](#-step-by-step-implementation)
  - [Step 1: Prepare the Google Sheet 📄](#step-1-prepare-the-google-sheet-)
  - [Step 2: Configure the Google Apps Script ⚙️](#step-2-configure-the-google-apps-script-️)
  - [Step 3: Set Up the Slack Bot 🤖](#step-3-set-up-the-slack-bot-)
  - [Step 4: Create the Slack Workflow 🌊](#step-4-create-the-slack-workflow-)
  - [Step 5: Final Authorization and Deployment 🎉](#step-5-final-authorization-and-deployment-)

---

## 📖 Introduction

This system uses a **Slack Workflow** to capture IT support requests triggered by an emoji reaction. These requests are automatically logged into a **Google Sheet**. A **Google Apps Script** then enriches this data by:

- Fetches the full text of a Slack message from a permalink.
- Determines the user's location based on the Slack channel ID.
- Categorizes the ticket by searching the message for keywords.
- Assigns a unique, formatted Ticket ID.
- Formats rows with colors based on location.
- Periodically removes duplicate entries.
- Sends a daily summary report of open tickets to a specified Slack channel.

---

## ✅ Prerequisites

Before you begin, ensure you have the following:

- **Admin access** to your company's Slack workspace.
- A **Google account** with access to Google Sheets and Google Apps Script.
- A **Google Sheet** to act as the ticketing database. You can create a new one or use the [provided template](https://docs.google.com/spreadsheets/d/1ZmXAURXe6KZyaIaFbbFxQaARn8-Weyg11CoTDvUJDj0/edit?usp=sharing).

---

## 🚀 Step-by-Step Implementation

### Step 1: Prepare the Google Sheet 📄

1.  **Create a Copy**: Make a copy of the [Ticketing System Template Google Sheet](https://docs.google.com/spreadsheets/d/1ZmXAURXe6KZyaIaFbbFxQaARn8-Weyg11CoTDvUJDj0/edit?usp=sharing).
2.  **Verify Columns**: Your sheet should have the following columns (the template is pre-configured):
    - `A: Timestamp`
    - `B: Sender`
    - `C: Message`
    - `D: Message Link`
    - `E: Location`
    - `F: Category`
    - `G: Status`
    - `H: Closed By`
    - `I: Closed Ticket Timestamp`
    - `J: Ticket ID`
    - `K: Solution`
    - `L: Notes`
3.  **Access Apps Script**: Open your new Google Sheet and navigate to **Extensions** > **Apps Script**.

### Step 2: Configure the Google Apps Script ⚙️

1.  **Paste the Code**: If you used the template, the code is already there. Otherwise, paste the provided code [TicketingScript.gs] into the script editor.
2.  **Set Script Properties (IMPORTANT)**:
    - In the Apps Script editor, go to **Project Settings** (the gear icon ⚙️ on the left).
    - Scroll down to **Script Properties** and click **Add script property**.
    > **Create a new property:**
    > - **Property**: `SLACK_BOT_TOKEN`
    > - **Value**: Paste your Slack Bot Token here (generated in the next step). It will start with `xoxb-`.
    - Click **Save script properties**.
3.  **Update Configuration Variables**: In the code itself, update the variables in the `--- Configuration ---` section:
    - `TARGET_SHEET_NAME`: The exact name of the sheet tab (e.g., `IT Tickets`).
    - `REPORTING_CHANNEL_ID`: The Slack Channel ID for daily summary reports.
    - `CHANNEL_TO_LOCATION_MAP`: Map your helpdesk Slack channel IDs to location names.

### Step 3: Set Up the Slack Bot 🤖

1.  **Create a Slack App**:
    - Go to `api.slack.com/apps` and click **Create New App** > **From scratch**.
    - Name your app (e.g., "IT Ticket Bot") and select your workspace.
2.  **Add Permissions**:
    - Go to the **OAuth & Permissions** page.
    - Under **Bot Token Scopes**, add the following permissions:
      - `channels:history`
      - `chat:write`
      - `reactions:read`
      - `users:read.email`
3.  **Install the App**:
    - Click **Install to Workspace** at the top of the **OAuth & Permissions** page and authorize it.
4.  **Get the Bot Token**:
    - After installation, copy the **Bot User OAuth Token** (starts with `xoxb-`).
    - Paste this token into the **Script Properties** in your Google Apps Script as described in Step 2.

### Step 4: Create the Slack Workflow 🌊

This single workflow handles the entire ticket lifecycle.

1.  **Open Workflow Builder**: In Slack, click your workspace name > **Tools** > **Workflow Builder**.
2.  **Create a New Workflow**: Click **Create** and name it "IT Ticketing System".
3.  **Configure Workflow Steps**:

    - #### **Trigger: Emoji Reaction**
      - **Trigger**: "When an emoji reaction is used".
      - **Channel(s)**: Your `#it-helpdesk` channel.
      - **Emoji**: A specific emoji like `:ticket:`.

    - #### **Action: Create Ticket in Sheet**
      - **Step**: Google Sheets > "Add a spreadsheet row".
      - **Map Data**:
        - `A: Timestamp`: Insert variable > Timestamp reacted.
        - `B: Sender`: Insert variable > Person who sent message > Email.
        - `D: Message Link`: Insert variable > Link to message.
        - `G: Status`: Type the word `Open`.

    - #### **Action: "Ticket Opened" Reply**
      - **Step**: "Send a message".
      - **Settings**: Check the box to **Reply in a thread** to the `Message that was reacted to`.
      - **Message**: `<@(Person who sent the message that was reacted to)> your ticket has been opened! @it has been notified and will begin investigating. 🧑‍🚀`
      - **Button**: Add a button with the text `Close Ticket`.

    - #### **Action: Update Sheet to "Closed"**
      - **Step**: Google Sheets > "Update a spreadsheet row".
      - **Find Row**: Search column `D: Message Link` for value `Link to message that was reacted to`.
      - **Update Columns**:
        - `G: Status`: Type `Closed`.
        - `H: Closed By`: Insert variable > Person who clicked the button > Email.
        - `I: Closed Ticket Timestamp`: Insert variable > Timestamp of button click.

    - #### **Action: "Ticket Closed" Reply**
      - **Step**: Threaded "Send a message".
      - **Message**: `<@(Person who sent the message that was reacted to)> this ticket has now been closed!`
      - **Button**: Add a button with the text `Re-Open Ticket`.

    - #### **Action: Update Sheet to "Open"**
      - **Step**: Google Sheets > "Update a spreadsheet row".
      - **Find Row**: Use the same search condition as before.
      - **Update Columns**:
        - `G: Status`: Type `Open`.
        - `H: Closed By`: Leave blank to clear.
        - `I: Closed Ticket Timestamp`: Leave blank to clear.

    - #### **Action: "Ticket Re-Opened" Reply**
      - **Step**: Threaded "Send a message".
      - **Message**: `<@(Person who sent the message that was reacted to)> this ticket has re-opened and @it has been notified!`
      - **Button**: Add the `Close Ticket` button again.

    - #### **Final Actions: Close Again**
      - Repeat the **Update Sheet to "Closed"** and **"Ticket Closed" Reply** steps.
      - **Do not** add any buttons to the final "Ticket Closed" message.

4.  **Publish** the workflow.

### Step 5: Final Authorization and Deployment 🎉

1.  **Save the Script**: In the Apps Script editor, click the **Save project** icon (💾).
2.  **Run the Trigger Setup**:
    - From the function dropdown list, select `createOrReplaceDailyTriggers` and click **Run**.
    > **Authorization Required**
    > A popup will appear asking for authorization.
    > 1. Click **Review permissions** and choose your Google account.
    > 2. You may see a "Google hasn't verified this app" screen. Click **Advanced** and then **Go to (your script name) (unsafe)**.
    > 3. Review the permissions and click **Allow**.
3.  **Verify Triggers**: Go to the **Triggers** page (the clock icon ⏰ on the left). You should see two new triggers: `removeDuplicatesAndProcess` and `sendDailyStatusReport`.

Your system is now live! Test it by reacting to a message in your helpdesk channel with your chosen emoji.
