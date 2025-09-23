/**
 * GitHub Repository Tracker with Canvas Integration for Google Sheets
 * This script fetches student submissions from Canvas and checks GitHub repositories
 * 
 * Setup Instructions:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code and save
 * 4. Set up your GitHub token and Canvas API token in script properties
 * 5. Configure your Canvas course and assignment IDs
 * 6. Run the setup function first
 */

// Configuration - Update these values
const CANVAS_BASE_URL = 'https://byupw.instructure.com'; // Your Canvas instance URL
const COURSE_ID = '12345'; // Your Canvas course ID
const ASSIGNMENT_ID = '56789'; // The specific assignment ID to check
const ASSIGNMENT_ID_FALLBACK = '67890'; // Fallback assignment ID to check if no submission or username
const REPO_NAME = 'wdd130'; // Main repository name to check
const REPO_VARIATIONS = ['wdd-130', 'WDD-130', 'wdd_130', 'WDD130', 'WD130', 'wd130', 'Wd130', 'Wdd130']; // Alternative names to check
// CAA (Code-Along Activity)
const CAA_CHECKS = [
  { 
    column: 'Week 01 CAA', 
    type: 'file_exists', 
    path: 'week01/favorite-city.html',
    description: 'favorite-city.html exists'
  },
  { 
    column: 'Week 02 CAA', 
    type: 'file_exists', 
    path: 'week02/styles/temple.css',
    description: 'temple.css exists'
  },
  { 
    column: 'Week 03 CAA', 
    type: 'code_search', 
    file: 'index.html', 
    searchFor: 'class="box"',
    description: 'class="box" exists in index.html'
  },
  {
    column: 'About Us Page',
    type: 'code_search',
    file: 'wwr/about.html',
    searchFor: 'meta name="description"',
    description: 'About Us Page exists'
  },
  { 
    column: 'Week 04 CAA', 
    type: 'code_search', 
    file: 'styles/styles.css', 
    searchFor: 'space-evenly',
    description: 'space-evenly exists in styles/styles.css'
  },
  { 
    column: 'Week 05 CAA', 
    type: 'file_exists', 
    path: 'week05/quiz.html',
    description: 'quiz.html exists'
  },
  {
    column: 'Contact Us Page',
    type: 'code_search',
    file: 'wwr/contact.html',
    searchFor: 'iframe',
    description: 'Contact Us Page exists'
  },
  {
    column: 'Home Page',
    type: 'code_search',
    file: 'wwr/index.html',
    searchFor: 'name="description" content="',
    description: 'Home Page exists'
  },
  {
    column: 'Trips Page',
    type: 'code_search',
    file: 'wwr/trips.html',
    searchFor: 'name="description" content="',
    description: 'Trips Page exists'
  }
];

// Basic checks for core files and code
const BASIC_CHECKS = {
  files: ['index.html'],
  code: [
    { file: 'index.html', searchFor: '<title>', description: 'Has title tag' },
    { file: 'index.html', searchFor: 'viewport', description: 'Has viewport' }
  ]
};

// Add types for Canvas student and submission objects
type CanvasStudent = {
  id: number;
  name: string;
  // ...other properties if needed
};

type CanvasSubmission = {
  user_id: number;
  workflow_state?: string;
  submitted_at?: string;
  url?: string;
  body?: string;
  // ...other properties if needed
};

/**
 * Setup function - Run this once to initialize
 */
function setup() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Define exact column headers as requested
  const headers = [
    'Student Name',          // A
    'Canvas User ID',        // B
    'GitHub Username',       // C
    'Repo URL',              // D
    'Repo Found',            // E
    'Alternative Found',     // F
    'Alt Repo URL',          // G
    'Has index.html',        // H
    'Has title tag',         // I
    'Has viewport',          // J
    'Week 01 CAA',           // K
    'Week 02 CAA',           // L
    'Week 03 CAA',           // M
    'About Us Page',         // N
    'Week 04 CAA',           // O
    'Week 05 CAA',           // P
    'Contact Us Page',       // Q
    'Home Page',             // R
    'Trips Page',            // S
    'Last Updated',          // T
    'Notes'                  // U
  ];
  
  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);

  // Define separate ranges to avoid overlap
  var range = sheet.getRange('E2:S100');   // Column E (Repo Found)

  // Rules for E:S - "Yes" = green, "No" = red
  var yesRuleNormal = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Yes')
      .setBackground('#CBEBCB') // Light green
      .setFontColor('#205520') // Dark green
      .setBold(true)
      .setRanges([range])  // Apply to both ranges
      .build();

  var noRuleNormal = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('No')
      .setBackground('#FCC6BB') // Light red
      .setFontColor('#9C2007') // Dark red
      .setBold(true)
      .setRanges([range])  // Apply to both ranges
      .build();

  // Clear existing rules and apply new ones
  sheet.clearConditionalFormatRules();
  sheet.setConditionalFormatRules([yesRuleNormal, noRuleNormal]);

  // Set text alignment to center for the entire range
  range.setHorizontalAlignment('center');
    
  Logger.log('Setup complete! Column layout updated.');
  Logger.log('Required properties:');
  Logger.log('- GITHUB_TOKEN: your_github_personal_access_token');
  Logger.log('- CANVAS_TOKEN: your_canvas_api_token');
}

/**
 * Fetch all students and their submissions, populate spreadsheet
 */
function fetchCanvasSubmissions(): void {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Get all students in the course
    const students: CanvasStudent[] = getAllStudents();
    // Get submissions for the main assignment
    const submissionsMain: CanvasSubmission[] = getCanvasSubmissions(ASSIGNMENT_ID);
    // Get submissions for the fallback assignment
    const submissionsFallback: CanvasSubmission[] = getCanvasSubmissions(ASSIGNMENT_ID_FALLBACK);
    
    // Create maps of submissions by user ID for easy lookup
    const submissionMapMain: { [userId: number]: CanvasSubmission } = {};
    submissionsMain.forEach(sub => {
      submissionMapMain[sub.user_id] = sub;
    });
    const submissionMapFallback: { [userId: number]: CanvasSubmission } = {};
    submissionsFallback.forEach(sub => {
      submissionMapFallback[sub.user_id] = sub;
    });
    
    // Clear existing data (keep headers)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    
    // Process each student
    const studentData = [];
    
    for (const student of students) {
      const studentName = student.name;
      const userId = student.id;
      let submission = submissionMapMain[userId];
      let submittedAt = '';
      let githubUsername = '';
      let submissionUrl = '';
      let submissionStatus = 'Not Submitted';

      // Try main assignment first
      if (submission) {
        submissionStatus = submission.workflow_state || 'Submitted';
        submittedAt = submission.submitted_at ? new Date(submission.submitted_at).toISOString() : '';
        if (submission.url) {
          submissionUrl = submission.url;
          githubUsername = extractGithubUsername(submission.url);
        } else if (submission.body && submission.body.includes('github.io')) {
          const urlMatch = submission.body.match(/https?:\/\/([^\.]+)\.github\.io[^\s]*/);
          if (urlMatch) {
            submissionUrl = urlMatch[0];
            githubUsername = extractGithubUsername(submissionUrl);
          }
        }
      }

      // If no submission or no GitHub username, try fallback assignment
      if (!submission || !githubUsername) {
        const fallbackSubmission = submissionMapFallback[userId];
        if (fallbackSubmission) {
          submissionStatus = fallbackSubmission.workflow_state || 'Submitted';
          submittedAt = fallbackSubmission.submitted_at ? new Date(fallbackSubmission.submitted_at).toISOString() : '';
          if (fallbackSubmission.url) {
            submissionUrl = fallbackSubmission.url;
            githubUsername = extractGithubUsername(fallbackSubmission.url);
          } else if (fallbackSubmission.body && fallbackSubmission.body.includes('github.io')) {
            const urlMatch = fallbackSubmission.body.match(/https?:\/\/([^\.]+)\.github\.io[^\s]*/);
            if (urlMatch) {
              submissionUrl = urlMatch[0];
              githubUsername = extractGithubUsername(submissionUrl);
            }
          }
        }
      }
      
      studentData.push([
        studentName,          // A - Student Name
        userId,               // B - Canvas User ID
        githubUsername,       // C - GitHub Username
        submissionUrl,        // D - Repo URL (submission URL for now)
        '',                   // E - Repo Found (will be filled by checkAllRepositories)
        '',                   // F - Alternative Found
        '',                   // G - Alt Repo URL
        '',                   // H - Has index.html
        '',                   // I - Has title tag
        '',                   // J - Has viewport
        '',                   // K - Week 01 CAA
        '',                   // L - Week 02 CAA
        '',                   // M - Week 03 CAA
        '',                   // N - About Us Page
        '',                   // O - Week 04 CAA
        '',                   // P - Week 05 CAA
        '',                   // Q - Contact Us Page
        '',                   // R - Home Page
        '',                   // S - Trips Page
        '',             // T - Last Updated
        submissionStatus === 'Not Submitted' ? 'No submission found' : '' // U - Notes
      ]);
    }
    
    if (studentData.length > 0) {
      sheet.getRange(2, 1, studentData.length, studentData[0].length).setValues(studentData);
      
      const submittedCount = studentData.filter(row => row[4] !== 'Not Submitted').length;
      const notSubmittedCount = studentData.length - submittedCount;
      
      Logger.log(`Imported ${studentData.length} students from Canvas`);
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, studentData[0].length);
      
      SpreadsheetApp.getUi().alert(
        `Success!\n\nImported ${studentData.length} students from Canvas:\n` +
        `• ${submittedCount} submitted\n• ${notSubmittedCount} not submitted\n\n` +
        `Click "GitHub Tracker" → "Check All Repositories" to analyze GitHub repos.`
      );
    } else {
      SpreadsheetApp.getUi().alert('No students found in this course.');
    }
    
  } catch (error) {
    Logger.log(`Error fetching Canvas data: ${error.toString()}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.toString()}`);
  }
}

/**
 * Get all students enrolled in the course
 */
function getAllStudents(): CanvasStudent[] {
  const token = PropertiesService.getScriptProperties().getProperty('CANVAS_TOKEN');
  
  if (!token) {
    throw new Error('Canvas API token not found. Please add CANVAS_TOKEN to Script Properties.');
  }
  
  const url = `${CANVAS_BASE_URL}/api/v1/courses/${COURSE_ID}/users?enrollment_type[]=student&per_page=100`;
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  };
  
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Canvas API error: ${response.getResponseCode()} - ${response.getContentText()}`);
  }
  
  return JSON.parse(response.getContentText()) as CanvasStudent[];
}

/**
 * Get submissions from Canvas API
 * Accepts assignmentId as parameter
 */
function getCanvasSubmissions(assignmentId: string): CanvasSubmission[] {
  const token = PropertiesService.getScriptProperties().getProperty('CANVAS_TOKEN');
  
  if (!token) {
    throw new Error('Canvas API token not found. Please add CANVAS_TOKEN to Script Properties.');
  }
  
  const url = `${CANVAS_BASE_URL}/api/v1/courses/${COURSE_ID}/assignments/${assignmentId}/submissions?per_page=100`;
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  };
  
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Canvas API error: ${response.getResponseCode()} - ${response.getContentText()}`);
  }
  
  return JSON.parse(response.getContentText()) as CanvasSubmission[];
}

/**
 * Extract GitHub username from various URL formats
 */
function extractGithubUsername(url: string): string {
  if (!url) return '';
  
  // Clean up the URL - remove extra whitespace
  url = url.trim();
  
  // Common GitHub Pages patterns:
  // https://username.github.io/repo-name
  // https://username.github.io/repo-name/
  // https://username.github.io/
  // Also handle GitHub repo links: https://github.com/username/repo-name
  
  const patterns = [
    /https?:\/\/([^\.\/\s]+)\.github\.io/i,  // GitHub Pages (case insensitive)
    /github\.com\/([^\/\s]+)/i,              // GitHub repo links (case insensitive)
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1] && match[1] !== 'www') {
      // Convert to lowercase and remove any invalid characters
      let username = match[1].toLowerCase();
      // GitHub usernames can only contain alphanumeric characters and hyphens
      username = username.replace(/[^a-z0-9\-]/g, '');
      
      // Validate it looks like a reasonable GitHub username
      if (username.length > 0 && username.length <= 39 && !username.startsWith('-') && !username.endsWith('-')) {
        Logger.log(`Extracted GitHub username: "${username}" from URL: ${url}`);
        return username;
      }
    }
  }
  
  Logger.log(`Could not extract GitHub username from URL: ${url}`);
  return '';
}

/**
 * Get canonical GitHub username (case and spelling as stored on GitHub)
 */
function getCanonicalGithubUsername(username: string): string | null {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const url = `https://api.github.com/users/${username}`;
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `token ${token}`,
      'Accept': 'application/vnd.github.v3+json'
    },
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return null;
    const user = JSON.parse(response.getContentText());
    return user.login || null;
  } catch (error) {
    Logger.log(`Error fetching canonical username for ${username}: ${error.toString()}`);
    return null;
  }
}

/**
 * Check all repositories for a user and update spreadsheet
 */
function checkAllRepositories(): void {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const username = data[i][2];
    if(!username || username.trim() === '') {
      Logger.log(`${data[i][0]} does not have a GitHub Username`);
      continue;
    }
    checkUserRepository(i + 1, username);
    Utilities.sleep(200); // Rate limiting - wait a little bit between requests
  }
}

/**
 * Check a single user's repository using GraphQL after canonical username lookup
 */
function checkUserRepository(row: number, username: string): void {
  const sheet = SpreadsheetApp.getActiveSheet();
  try {
    // Step 1: Get canonical username
    const canonicalUsername = getCanonicalGithubUsername(username);
    if (!canonicalUsername) {
      Logger.log(`Could not find GitHub user for: ${username}`);
      updateRowWithResults(row, username, { exists: false, url: '', pagesEnabled: false, pagesUrl: '', basicChecks: {}, caaChecks: {} }, null);
      return;
    }

    // Step 2: Use GraphQL to check all repos/files
    const repoNames = [REPO_NAME, ...REPO_VARIATIONS];
    const filePaths = BASIC_CHECKS.files
      .concat(CAA_CHECKS.filter(c => c.type === 'file_exists').map(c => c.path))
      .concat(CAA_CHECKS.filter(c => c.type === 'code_search').map(c => c.file));
    const graphqlResult = githubGraphQLQuery(canonicalUsername, repoNames, filePaths);

    // Step 3: Parse GraphQL result and update spreadsheet
    let mainRepoData = { exists: false, url: '', pagesEnabled: false, pagesUrl: '', basicChecks: {}, caaChecks: {} };
    let altRepoData = null;

    if (graphqlResult && graphqlResult.data) {
      for (let i = 0; i < repoNames.length; i++) {
        const repoKey = `repo${i}`;
        const repo = graphqlResult.data[repoKey];
        if (repo) {
          const repoData = {
            exists: true,
            url: repo.url,
            pagesEnabled: false,
            pagesUrl: '',
            basicChecks: {},
            caaChecks: {}
          };

          // Basic file existence
          for (let f = 0; f < BASIC_CHECKS.files.length; f++) {
            const fileKey = `file${f}`;
            repoData.basicChecks[BASIC_CHECKS.files[f]] = !!repo[fileKey];
          }
          // Basic code checks
          for (const codeCheck of BASIC_CHECKS.code) {
            const fileIdx = filePaths.indexOf(codeCheck.file);
            if (fileIdx !== -1 && repo[`file${fileIdx}`] && repo[`file${fileIdx}`].text) {
              repoData.basicChecks[codeCheck.description] = repo[`file${fileIdx}`].text.toLowerCase().includes(codeCheck.searchFor.toLowerCase());
            } else {
              repoData.basicChecks[codeCheck.description] = false;
            }
          }
          // CAA checks
          for (const caa of CAA_CHECKS) {
            let idx = -1;
            if (caa.type === 'file_exists') {
              idx = filePaths.indexOf(caa.path);
              repoData.caaChecks[caa.column] = idx !== -1 && !!repo[`file${idx}`];
            } else if (caa.type === 'code_search') {
              idx = filePaths.indexOf(caa.file);
              repoData.caaChecks[caa.column] = idx !== -1 && repo[`file${idx}`] && repo[`file${idx}`].text
                ? repo[`file${idx}`].text.toLowerCase().includes(caa.searchFor.toLowerCase())
                : false;
            }
          }

          if (!mainRepoData.exists) {
            mainRepoData = repoData;
          } else {
            altRepoData = repoData;
          }
        }
      }
    }
    updateRowWithResults(row, canonicalUsername, mainRepoData, altRepoData);
  } catch (error) {
    Logger.log(`Error checking ${username}: ${error.toString()}`);
    sheet.getRange(row, sheet.getLastColumn()).setValue(`Error: ${error.toString()}`);
  }
}

/**
 * Custom menu for easy access
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GitHub Tracker')
    .addItem('1. Setup Spreadsheet', 'setup')
    .addSeparator()
    .addItem('2. Fetch Canvas Submissions', 'fetchCanvasSubmissions')
    .addItem('3. Check All Repositories', 'checkAllRepositories')
    .addSeparator()
    .addItem('Check Selected Row Only', 'checkSelectedRow')
    .addItem('Test Canvas Connection', 'testCanvasConnection')
    .addToUi();
}

/**
 * Test Canvas API connection
 */
function testCanvasConnection(): void {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('CANVAS_TOKEN');
    
    if (!token) {
      SpreadsheetApp.getUi().alert('Canvas API token not found. Please add CANVAS_TOKEN to Script Properties.');
      return;
    }
    
    // Test with course info
    const url = `${CANVAS_BASE_URL}/api/v1/courses/${COURSE_ID}`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const courseData = JSON.parse(response.getContentText());
      SpreadsheetApp.getUi().alert(`Success!\n\nConnected to Canvas course: "${courseData.name}"\nCourse ID: ${COURSE_ID}`);
    } else {
      SpreadsheetApp.getUi().alert(`Error: ${response.getResponseCode()} - ${response.getContentText()}`);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Connection failed: ${error.toString()}`);
  }
}

/**
 * Check repository for currently selected row
 */
function checkSelectedRow(): void {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Please select a row with student data (not the header row).');
    return;
  }
  
  const username = sheet.getRange(row, 3).getValue(); // Column C - GitHub Username
  
  if (!username || username.trim() === '') {
    SpreadsheetApp.getUi().alert('No GitHub username found in this row. Please fetch Canvas submissions first.');
    return;
  }
  
  checkUserRepository(row, username);
  SpreadsheetApp.getUi().alert('Repository check completed for ' + username);
}

/**
 * Query GitHub GraphQL API for repo existence and file content
 */
function githubGraphQLQuery(username: string, repoNames: string[], filePaths: string[]): any {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  // Build GraphQL query string
  let repoQueries = repoNames.map((repo, i) => `
    repo${i}: repository(owner: "${username}", name: "${repo}") {
      name
      url
      ${filePaths.map((file, j) => `
        file${j}: object(expression: "HEAD:${file}") {
          ... on Blob {
            text
          }
        }
      `).join('\n')}
    }
  `).join('\n');
  const query = `query {
    ${repoQueries}
  }`;
  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      'Authorization': `bearer ${token}`
    },
    payload: JSON.stringify({ query }),
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch('https://api.github.com/graphql', options);
  if (response.getResponseCode() !== 200) {
    Logger.log('GraphQL error: ' + response.getContentText());
    return null;
  }
  return JSON.parse(response.getContentText());
}

// Example usage inside your repo check logic:
// const result = githubGraphQLQuery(username, [REPO_NAME, ...REPO_VARIATIONS], ['index.html', 'README.md']);
// Then parse result.data.repo0, repo1, etc.