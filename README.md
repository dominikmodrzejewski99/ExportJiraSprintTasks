# Jira Sprint Report Generator

This package provides multiple scripts for generating Excel reports from Jira sprints. Each script offers different features to suit various needs.

## Available Scripts

The package includes several scripts with different features:

### 1. Simple Report (generate-sprint-report-auto.js)

- Includes all issues from the sprint without filtering
- No grouping or ordering by assignee
- Simplest approach with minimal configuration

### 2. Ordered Report (generate-sprint-report-ordered-env.js)

- Groups issues by assignee
- Orders developers in a specific sequence (configurable)
- Excludes issues assigned to testers (configurable)
- Provides the most structured output

## Setup

1. Copy the `.env.example` file to `.env`:
   ```
   cp .env.example .env
   ```

2. Edit the `.env` file and fill in your Jira credentials:
   ```
   JIRA_API_KEY=your_api_key
   JIRA_BASE_URL=https://your_company.atlassian.net
   JIRA_EMAIL=your.email@company.com
   DEFAULT_BOARD_ID=your_board_id
   ```

   You can generate a Jira API key at: https://id.atlassian.com/manage-profile/security/api-tokens

3. Install dependencies:
   ```
   npm install
   ```

## Usage

### Using npm Scripts

The easiest way to run the scripts is using the npm commands defined in package.json:

```bash
# Run the simple report
npm start

# Run the ordered report with developer grouping
npm run start-ordered
```

### Running Scripts Directly

#### Simple Report

Run the script without arguments to use the default board ID from your `.env` file:

```bash
node generate-sprint-report-auto.js
```

#### Ordered Report

```bash
node generate-sprint-report-ordered-env.js
```

### Specifying a Board

You can specify a board ID or name with either script:

```bash
# By ID
node generate-sprint-report-auto.js 125

# By name
node generate-sprint-report-auto.js "CMS"
```

The script will find all boards containing "CMS" in their name and use the first one.

## Configuration Options

### Basic Configuration

```
JIRA_API_KEY=your_api_key
JIRA_BASE_URL=https://your_company.atlassian.net
JIRA_EMAIL=your.email@company.com
DEFAULT_BOARD_ID=your_board_id
```

### Column Names

You can customize the column names in the Excel report:

```
COLUMN_ISSUE_KEY=Task ID
COLUMN_ISSUE_TYPE=Type
COLUMN_SUMMARY=Description
COLUMN_ASSIGNEE=Developer
COLUMN_STATUS=Current Status
```

### Developer Order (Ordered Report Only)

Specify the order in which developers should appear in the report:

```
DEVELOPER_ORDER=John Smith,Jane Doe,Bob Johnson
```

### Testers to Exclude (Ordered Report Only)

Specify which assignees should be excluded from the report:

```
TESTERS_TO_EXCLUDE=Tester One,Tester Two
```

## Automation

You can set up a cron job to run this script automatically on a regular basis:

```
# Run every Monday at 8 AM
0 8 * * 1 cd /path/to/script/directory && npm start
```

## Output

The script generates an Excel file named after the sprint, for example:
```
Sprint_5ways-jg-65_2025-04-08.xlsx
```

The file includes:
- Issues from the sprint (filtered according to the script used)
- Color coding based on status (green for Closed/Done, orange for In Progress, blue for Review)
- Formatted headers and borders
- Auto-filtering and frozen header row

## Choosing the Right Script

- **Simple Report**: Use when you want a complete list of all issues without any filtering or special ordering.
- **Ordered Report**: Use when you want issues grouped by developer in a specific order, with the ability to exclude certain assignees (like testers).
