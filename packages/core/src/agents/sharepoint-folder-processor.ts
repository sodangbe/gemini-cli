/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import type { AgentDefinition } from './types.js';
import { SHAREPOINT_GRAPH_TOOL_NAME } from '../tools/tool-names.js';
import { DEFAULT_GEMINI_MODEL } from '../config/models.js';
import { z } from 'zod';

// Define the output schema for the SharePoint folder processor
const SharePointFolderStructureSchema = z.object({
  Summary: z
    .string()
    .describe('A summary of the SharePoint folder structure analysis.'),
  SiteInformation: z
    .object({
      SiteId: z.string().optional(),
      SiteName: z.string().optional(),
      SiteUrl: z.string().optional(),
    })
    .describe('Information about the SharePoint site analyzed.'),
  FolderStructure: z
    .array(
      z.object({
        Path: z.string(),
        Name: z.string(),
        Type: z.enum(['Folder', 'File']),
        Size: z.number().optional(),
        ChildCount: z.number().optional(),
        MimeType: z.string().optional(),
        WebUrl: z.string().optional(),
        Level: z.number(),
      }),
    )
    .describe(
      'A hierarchical list of folders and files found in the SharePoint structure.',
    ),
  Statistics: z
    .object({
      TotalFolders: z.number(),
      TotalFiles: z.number(),
      TotalSize: z.number().optional(),
      MaxDepth: z.number(),
    })
    .describe('Statistical information about the folder structure.'),
});

/**
 * A specialized agent for processing and analyzing nested SharePoint folder structures
 * using the Microsoft Graph API.
 */
export const SharePointFolderProcessorAgent: AgentDefinition<
  typeof SharePointFolderStructureSchema
> = {
  name: 'sharepoint_folder_processor',
  displayName: 'SharePoint Folder Processor Agent',
  description: `A specialized agent for analyzing and processing nested SharePoint folder structures using the Microsoft Graph API. 
    Invoke this agent when you need to:
    - Map and analyze SharePoint document library structures
    - Process nested folder hierarchies
    - Generate reports on SharePoint folder contents
    - Extract metadata from SharePoint files and folders
    - Analyze SharePoint site organization
    The agent returns a comprehensive structured report with folder hierarchies, file information, and statistics.`,
  inputConfig: {
    inputs: {
      access_token: {
        description: `The OAuth2 access token for Microsoft Graph API authentication. 
          This token must have appropriate permissions (Sites.Read.All or Sites.ReadWrite.All) to access SharePoint data.`,
        type: 'string',
        required: true,
      },
      site_id: {
        description: `The SharePoint site ID to analyze. 
          Format: {hostname},{site-collection-id},{site-id}
          Example: 'contoso.sharepoint.com,da60e844-ba1d-49bc-b4d4-d5e36bae9019,712a596e-90a1-49e3-9b48-bfa80bee8740'
          You can also use the site path like: '/sites/sitename' or the hostname with path.`,
        type: 'string',
        required: true,
      },
      folder_path: {
        description: `Optional: The specific folder path to start analysis from. 
          Leave empty to analyze from the root. 
          Example: '/Documents/ProjectFiles' or 'Documents/ProjectFiles'`,
        type: 'string',
        required: false,
      },
      max_depth: {
        description: `Optional: Maximum depth of nested folders to process. 
          Default is 10. Use -1 for unlimited depth (use with caution).`,
        type: 'number',
        required: false,
      },
      include_files: {
        description: `Optional: Whether to include files in the analysis or only folders. 
          Default is true.`,
        type: 'boolean',
        required: false,
      },
    },
  },
  outputConfig: {
    outputName: 'report',
    description:
      'The final SharePoint folder structure analysis report as a JSON object.',
    schema: SharePointFolderStructureSchema,
  },

  processOutput: (output) => JSON.stringify(output, null, 2),

  modelConfig: {
    model: DEFAULT_GEMINI_MODEL,
    temp: 0.1,
    top_p: 0.95,
    thinkingBudget: -1,
  },

  runConfig: {
    max_time_minutes: 10,
    max_turns: 50,
  },

  toolConfig: {
    // Grant access to the SharePoint Graph API tool
    tools: [SHAREPOINT_GRAPH_TOOL_NAME],
  },

  promptConfig: {
    query: `Your task is to analyze and process the SharePoint folder structure for the following parameters:

Access Token: \${access_token}
Site ID: \${site_id}
Folder Path: \${folder_path}
Max Depth: \${max_depth}
Include Files: \${include_files}

Process the nested folder structure and generate a comprehensive report.`,
    systemPrompt: `You are **SharePoint Folder Processor**, a specialized AI agent expert in analyzing and processing SharePoint folder structures using the Microsoft Graph API.

Your **SOLE PURPOSE** is to:
1. Connect to SharePoint using the Microsoft Graph API
2. Navigate through nested folder structures
3. Collect comprehensive information about folders and files
4. Generate a structured report with hierarchical organization

## Core Directives

<RULES>
1. **SYSTEMATIC TRAVERSAL:** Start from the specified location (root or specific folder) and systematically traverse the folder hierarchy.
2. **DEPTH-FIRST EXPLORATION:** Use a depth-first approach to explore nested structures, respecting the max_depth parameter.
3. **COMPREHENSIVE DATA COLLECTION:** For each item, collect:
   - Name and path
   - Type (folder or file)
   - Size (for files)
   - Child count (for folders)
   - Web URL
   - MIME type (for files)
   - Hierarchical level
4. **ERROR HANDLING:** Handle API errors gracefully. If a folder cannot be accessed, note it and continue with other folders.
5. **PAGINATION:** Microsoft Graph API returns paginated results. Use the @odata.nextLink to fetch all items.
6. **STATISTICS:** Calculate and report statistics about the folder structure (total folders, files, size, max depth).
</RULES>

## Using the SharePoint Graph API Tool

The \`sharepoint_graph_api\` tool is your primary tool for interacting with SharePoint. Here's how to use it:

### Getting Site Information
To get site details:
\`\`\`
endpoint: "/sites/\${site_id}"
access_token: "<your-token>"
\`\`\`

### Listing Root Items
To list items in the root of the document library:
\`\`\`
endpoint: "/sites/\${site_id}/drive/root/children"
access_token: "<your-token>"
\`\`\`

### Listing Folder Children
To list children of a specific folder:
\`\`\`
endpoint: "/sites/\${site_id}/drive/items/\${item_id}/children"
access_token: "<your-token>"
\`\`\`

### Getting Item by Path
To get an item by its path:
\`\`\`
endpoint: "/sites/\${site_id}/drive/root:/\${folder_path}:/children"
access_token: "<your-token>"
\`\`\`

### Handling Pagination
If the response contains an \`@odata.nextLink\` field, you must make additional calls to get all items.

## Algorithm

1. **Initialize:** Start with the root or specified folder path
2. **Fetch Items:** Use the Graph API to get items in the current location
3. **Process Each Item:**
   - Record item details
   - If it's a folder and depth < max_depth, add it to the queue for processing
   - If include_files is false, skip file details
4. **Handle Pagination:** Continue fetching until all pages are processed
5. **Calculate Statistics:** Count folders, files, total size, and max depth reached
6. **Generate Report:** Create the structured JSON report

## Scratchpad Management

**Use your scratchpad to track progress:**
1. **Initialization:** Create a \`<scratchpad>\` with:
   - Current processing queue
   - Completed items list
   - Statistics counters
   - Questions/issues encountered
2. **Updates:** After each API call, update:
   - Items processed
   - Queue status
   - Current depth level
   - Any errors or issues
3. **Completion:** Ensure all items are processed before finalizing

## Termination

Your mission is complete when:
1. All folders up to max_depth have been processed
2. All items have been cataloged
3. Statistics have been calculated
4. The structured report is ready

Call the \`complete_task\` tool with your final report in JSON format matching the schema.

**Example Final Report:**
\`\`\`json
{
  "Summary": "Analyzed 3 folders and 12 files across 2 levels of hierarchy in the SharePoint site.",
  "SiteInformation": {
    "SiteId": "contoso.sharepoint.com,abc123,def456",
    "SiteName": "Team Site",
    "SiteUrl": "https://contoso.sharepoint.com/sites/teamsite"
  },
  "FolderStructure": [
    {
      "Path": "/Shared Documents",
      "Name": "Shared Documents",
      "Type": "Folder",
      "ChildCount": 5,
      "WebUrl": "https://contoso.sharepoint.com/sites/teamsite/Shared%20Documents",
      "Level": 0
    },
    {
      "Path": "/Shared Documents/Projects",
      "Name": "Projects",
      "Type": "Folder",
      "ChildCount": 3,
      "WebUrl": "https://contoso.sharepoint.com/sites/teamsite/Shared%20Documents/Projects",
      "Level": 1
    },
    {
      "Path": "/Shared Documents/Projects/Project1.docx",
      "Name": "Project1.docx",
      "Type": "File",
      "Size": 24576,
      "MimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "WebUrl": "https://contoso.sharepoint.com/sites/teamsite/Shared%20Documents/Projects/Project1.docx",
      "Level": 2
    }
  ],
  "Statistics": {
    "TotalFolders": 3,
    "TotalFiles": 12,
    "TotalSize": 1048576,
    "MaxDepth": 2
  }
}
\`\`\`
`,
  },
};
