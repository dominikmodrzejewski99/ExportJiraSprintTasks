// Load environment variables from .env file
require('dotenv').config();

// Import required modules
const JiraClient = require('jira-client');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Load config from environment variables
const jiraApiKey = process.env.JIRA_API_KEY;
const jiraBaseUrl = process.env.JIRA_BASE_URL;
const jiraEmail = process.env.JIRA_EMAIL;

// Default board ID from environment variables or fallback to 125 (CMS board)
const DEFAULT_BOARD_ID = process.env.DEFAULT_BOARD_ID || 125;

// Custom column names from environment variables or defaults
const COLUMN_NAMES = {
  ISSUE_KEY: process.env.COLUMN_ISSUE_KEY || 'Issue key',
  ISSUE_TYPE: process.env.COLUMN_ISSUE_TYPE || 'Issue Type',
  SUMMARY: process.env.COLUMN_SUMMARY || 'Summary',
  ASSIGNEE: process.env.COLUMN_ASSIGNEE || 'Assignee',
  STATUS: process.env.COLUMN_STATUS || 'Status'
};

// Parse testers list from environment variable
const TESTERS = process.env.TESTERS_TO_EXCLUDE
  ? process.env.TESTERS_TO_EXCLUDE.split(',').map(name => name.trim())
  : [];

// Parse developer order from environment variable
const DEVELOPER_ORDER = process.env.DEVELOPER_ORDER
  ? process.env.DEVELOPER_ORDER.split(',').map(name => name.trim())
  : [];

console.log('Jira API Key:', jiraApiKey ? 'API Key is set' : 'API Key is not set');
console.log('Jira Base URL:', jiraBaseUrl);
console.log('Jira Email:', jiraEmail);
console.log('Default Board ID:', DEFAULT_BOARD_ID);
console.log('Testers to exclude:', TESTERS.join(', '));
console.log('Developer order:', DEVELOPER_ORDER.join(', '));

// Initialize Jira client
const jira = new JiraClient({
  protocol: 'https',
  host: jiraBaseUrl.replace('https://', ''),
  username: jiraEmail,
  password: jiraApiKey,
  apiVersion: '2',
  strictSSL: true
});

// Function to get all boards
async function getAllBoards() {
  try {
    console.log('Fetching all boards...');
    const boards = await jira.getAllBoards();
    console.log(`Found ${boards.values.length} boards`);
    return boards.values;
  } catch (error) {
    console.error('Error fetching boards:', error.message);
    throw error;
  }
}

// Function to get the current active sprint for a board
async function getCurrentSprint(boardId) {
  try {
    console.log(`Fetching active sprint for board ${boardId}...`);
    const sprints = await jira.getAllSprints(boardId, 0, 50, 'active');

    if (sprints && sprints.values && sprints.values.length > 0) {
      console.log(`Found active sprint: ${sprints.values[0].name} (ID: ${sprints.values[0].id})`);
      return sprints.values[0];
    } else {
      console.log('No active sprint found. Looking for most recent sprint...');
      // If no active sprint, try to get the most recent sprint
      const allSprints = await jira.getAllSprints(boardId, 0, 50);
      if (allSprints && allSprints.values && allSprints.values.length > 0) {
        // Sort by start date descending
        allSprints.values.sort((a, b) => new Date(b.startDate) - new Date(a.startDate));
        console.log(`Found most recent sprint: ${allSprints.values[0].name} (ID: ${allSprints.values[0].id})`);
        return allSprints.values[0];
      }
    }

    throw new Error('No sprints found for this board');
  } catch (error) {
    console.error('Error fetching current sprint:', error.message);
    throw error;
  }
}

// Function to get issues for the current sprint using JQL
async function getSprintIssuesByJQL(sprintId, maxResults = 500) {
  try {
    console.log(`Fetching issues for sprint ${sprintId} using JQL...`);
    const jql = `sprint = ${sprintId} ORDER BY assignee ASC, updated DESC`;
    const issues = await jira.searchJira(jql, { maxResults });
    console.log(`Found ${issues.issues.length} issues in sprint ${sprintId}`);
    return issues.issues;
  } catch (error) {
    console.error(`Error fetching issues for sprint ${sprintId}:`, error.message);
    throw error;
  }
}

// Function to group issues by assignee, excluding testers
function groupIssuesByAssignee(issues) {
  const groupedIssues = {};
  const assignees = new Set();

  // First pass: collect all assignees and initialize groups
  issues.forEach(issue => {
    if (!issue.fields.assignee) return;

    const assignee = issue.fields.assignee.displayName;

    // Skip issues assigned to testers
    if (TESTERS.includes(assignee)) {
      console.log(`Skipping issue ${issue.key} assigned to tester ${assignee}`);
      return;
    }

    assignees.add(assignee);
  });

  // If we have a developer order, use it to filter and order assignees
  let orderedAssignees = [];
  if (DEVELOPER_ORDER.length > 0) {
    // Only include developers that are in our order list
    orderedAssignees = DEVELOPER_ORDER.filter(dev => assignees.has(dev));

    // Initialize groups for ordered developers
    orderedAssignees.forEach(developer => {
      groupedIssues[developer] = [];
    });
  } else {
    // No order specified, use all assignees sorted alphabetically
    orderedAssignees = Array.from(assignees).sort();

    // Initialize groups for all assignees
    orderedAssignees.forEach(assignee => {
      groupedIssues[assignee] = [];
    });
  }

  // Second pass: assign issues to groups
  issues.forEach(issue => {
    if (!issue.fields.assignee) return;

    const assignee = issue.fields.assignee.displayName;

    // Skip issues assigned to testers
    if (TESTERS.includes(assignee)) return;

    // Skip if assignee is not in our list
    if (!orderedAssignees.includes(assignee)) {
      console.log(`Skipping issue ${issue.key} assigned to ${assignee} (not in ordered list)`);
      return;
    }

    groupedIssues[assignee].push(issue);
  });

  return { groupedIssues, orderedAssignees };
}

// Function to find a board by name
async function findBoardByName(boardName) {
  try {
    console.log(`Looking for board with name containing "${boardName}"...`);
    const boards = await getAllBoards();

    // Find boards that match the name (case-insensitive)
    const matchingBoards = boards.filter(board =>
      board.name.toLowerCase().includes(boardName.toLowerCase())
    );

    if (matchingBoards.length > 0) {
      console.log(`Found ${matchingBoards.length} matching boards:`);
      matchingBoards.forEach(board => {
        console.log(`- ${board.name} (ID: ${board.id})`);
      });

      // Return the first matching board
      return matchingBoards[0];
    } else {
      console.log(`No boards found with name containing "${boardName}"`);
      return null;
    }
  } catch (error) {
    console.error('Error finding board by name:', error.message);
    throw error;
  }
}

// Function to generate Excel report for sprint
async function generateSprintReport(boardId) {
  try {
    console.log(`Generating sprint report for board ${boardId}...`);

    // Get the current sprint
    const sprint = await getCurrentSprint(boardId);

    // Get issues for the sprint using JQL
    const issues = await getSprintIssuesByJQL(sprint.id);

    // Group issues by assignee, excluding testers
    const { groupedIssues, orderedAssignees } = groupIssuesByAssignee(issues);

    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Jira Sprint Report Generator';
    workbook.lastModifiedBy = 'Jira Sprint Report Generator';
    workbook.created = new Date();
    workbook.modified = new Date();

    // Add a worksheet
    const worksheet = workbook.addWorksheet(`${sprint.name}`);

    // Define columns with custom names from environment variables
    worksheet.columns = [
      { header: 'Programista', key: 'developer', width: 25 },
      { header: COLUMN_NAMES.ISSUE_KEY, key: 'key', width: 15 },
      { header: COLUMN_NAMES.ISSUE_TYPE, key: 'type', width: 15 },
      { header: COLUMN_NAMES.SUMMARY, key: 'summary', width: 50 },
      { header: COLUMN_NAMES.ASSIGNEE, key: 'assignee', width: 20 },
      { header: COLUMN_NAMES.STATUS, key: 'status', width: 15 }
    ];

    // Style the header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E0E0' } // Light gray background
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    // Add issues to the worksheet
    let rowIndex = 2; // Start from row 2 (after header)

    // Process each assignee group in the specified order
    for (const assignee of orderedAssignees) {
      const assigneeIssues = groupedIssues[assignee];

      if (assigneeIssues.length === 0) {
        console.log(`No issues found for ${assignee}, skipping...`);
        continue;
      }

      console.log(`Adding ${assigneeIssues.length} issues for ${assignee}...`);

      // Add each issue to the worksheet
      for (const issue of assigneeIssues) {
        const row = {
          developer: assignee, // Use full name
          key: issue.key,
          type: issue.fields.issuetype.name,
          summary: issue.fields.summary,
          assignee: assignee,
          status: issue.fields.status.name
        };

        worksheet.addRow(row);

        // Style the row based on status
        const currentRow = worksheet.getRow(rowIndex);

        // Apply different colors based on status
        if (row.status === 'Closed' || row.status === 'Done') {
          currentRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE6F0E6' } // Light green
          };
        } else if (row.status === 'In Dev' || row.status === 'In Progress') {
          currentRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFF0E0' } // Light orange
          };
        } else if (row.status === 'Review') {
          currentRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0F0FF' } // Light blue
          };
        }

        rowIndex++;
      }

      // Add a blank row between assignees
      worksheet.addRow({});
      rowIndex++;
    }

    // Add borders to all cells with data
    for (let i = 1; i < rowIndex; i++) {
      const row = worksheet.getRow(i);
      row.eachCell({ includeEmpty: false }, cell => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    }

    // Auto filter for the header row
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 6 }
    };

    // Freeze the header row
    worksheet.views = [
      { state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'A2' }
    ];

    // Write to file
    const fileName = `Sprint_${sprint.name.replace(/\s+/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`;
    await workbook.xlsx.writeFile(fileName);

    console.log(`Excel report generated successfully: ${fileName}`);
    return fileName;
  } catch (error) {
    console.error('Error generating sprint report:', error.message);
    throw error;
  }
}

// Main function
async function main() {
  try {
    let boardId;

    // Check if a board name was provided as an argument
    if (process.argv[2] && isNaN(process.argv[2])) {
      // If the argument is not a number, treat it as a board name
      const boardName = process.argv[2];
      const board = await findBoardByName(boardName);

      if (board) {
        boardId = board.id;
      } else {
        console.log(`Using default board ID: ${DEFAULT_BOARD_ID}`);
        boardId = DEFAULT_BOARD_ID;
      }
    } else {
      // If a numeric board ID was provided, use it
      boardId = process.argv[2] ? parseInt(process.argv[2]) : DEFAULT_BOARD_ID;
    }

    // Generate sprint report
    await generateSprintReport(boardId);

    console.log('Report generation completed successfully!');
  } catch (error) {
    console.error('Error in main function:', error.message);
    process.exit(1);
  }
}

// Run the main function
main();
