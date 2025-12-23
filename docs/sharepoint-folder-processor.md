# SharePoint Folder Processor Agent

A specialized agent for analyzing and processing nested SharePoint folder
structures using the Microsoft Graph API.

## Overview

The SharePoint Folder Processor Agent provides automated analysis of SharePoint
document libraries and folder hierarchies. It can traverse nested folder
structures, collect metadata about files and folders, and generate comprehensive
reports.

## Features

- **Nested Folder Traversal**: Automatically navigate through complex folder
  hierarchies
- **Microsoft Graph API Integration**: Uses Microsoft Graph API for secure
  access to SharePoint data
- **Structured Output**: Returns JSON-formatted reports with hierarchical
  information
- **Configurable Depth**: Control how deep to traverse into nested folders
- **Statistics**: Calculate total files, folders, sizes, and maximum depth
- **Pagination Support**: Handles large folder structures with automatic
  pagination

## Configuration

Enable the SharePoint Folder Processor Agent in your Gemini CLI configuration:

```typescript
const config = new Config({
  // ... other config options
  sharepointFolderProcessorSettings: {
    enabled: true,
    maxNumTurns: 50, // Maximum agent turns
    maxTimeMinutes: 10, // Maximum execution time
    thinkingBudget: -1, // Thinking token budget (-1 for unlimited)
    model: 'gemini-2.0-flash-exp', // Model to use
  },
});
```

## Usage

### Prerequisites

1. **Microsoft 365 Account**: You need access to a SharePoint site
2. **Access Token**: Obtain an OAuth2 access token with appropriate permissions:
   - `Sites.Read.All` (read-only access)
   - `Sites.ReadWrite.All` (read-write access)

### Getting an Access Token

You can obtain an access token using Azure AD authentication:

```javascript
// Example using @azure/msal-node
const msal = require('@azure/msal-node');

const config = {
  auth: {
    clientId: 'YOUR_CLIENT_ID',
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID',
    clientSecret: 'YOUR_CLIENT_SECRET',
  },
};

const cca = new msal.ConfidentialClientApplication(config);

const tokenRequest = {
  scopes: ['https://graph.microsoft.com/.default'],
};

const response = await cca.acquireTokenByClientCredential(tokenRequest);
const accessToken = response.accessToken;
```

### Invoking the Agent

Once configured, you can invoke the agent as a tool from the main Gemini CLI:

```typescript
// The agent is automatically registered as a tool when enabled
const result = await geminiClient.generateContent({
  tools: ['sharepoint_folder_processor'],
  prompt: `Analyze the SharePoint folder structure at site ID: contoso.sharepoint.com,abc123,def456 
           with access token: ${accessToken}`,
});
```

### Agent Parameters

The SharePoint Folder Processor Agent accepts the following parameters:

- **access_token** (required): OAuth2 access token for Microsoft Graph API
- **site_id** (required): SharePoint site identifier
  - Format: `{hostname},{site-collection-id},{site-id}`
  - Example:
    `contoso.sharepoint.com,da60e844-ba1d-49bc-b4d4-d5e36bae9019,712a596e-90a1-49e3-9b48-bfa80bee8740`
  - Alternative: Use site path like `/sites/sitename`
- **folder_path** (optional): Specific folder path to start from (default: root)
  - Example: `/Documents/ProjectFiles`
- **max_depth** (optional): Maximum folder depth to traverse (default: 10, -1
  for unlimited)
- **include_files** (optional): Whether to include files in the analysis
  (default: true)

### Example Output

The agent returns a structured JSON report:

```json
{
  "Summary": "Analyzed 5 folders and 23 files across 3 levels of hierarchy in the SharePoint site.",
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
    "TotalFolders": 5,
    "TotalFiles": 23,
    "TotalSize": 2048576,
    "MaxDepth": 3
  }
}
```

## SharePoint Graph API Tool

The agent uses the `sharepoint_graph_api` tool internally to make Microsoft
Graph API calls. This tool can also be used independently:

### Tool Parameters

- **endpoint** (required): Graph API endpoint path
  - Example: `/sites/{site-id}/drive/root/children`
- **access_token** (required): OAuth2 access token
- **method** (optional): HTTP method (GET, POST, PATCH, DELETE)
- **body** (optional): JSON body for POST/PATCH requests

### Common Endpoints

- List site root items: `/sites/{site-id}/drive/root/children`
- List folder children: `/sites/{site-id}/drive/items/{item-id}/children`
- Get item by path: `/sites/{site-id}/drive/root:/{path}:/children`
- Get site info: `/sites/{site-id}`

## Error Handling

The agent handles various error scenarios:

- **Invalid Credentials**: Returns error if access token is invalid or expired
- **Insufficient Permissions**: Returns error if token lacks required
  permissions
- **Resource Not Found**: Handles cases where folders or sites don't exist
- **Network Errors**: Gracefully handles timeout and connection issues
- **Pagination**: Automatically follows `@odata.nextLink` for large result sets

## Limitations

- Maximum execution time is configurable (default: 10 minutes)
- Maximum agent turns is configurable (default: 50 turns)
- Large folder structures may require multiple API calls due to pagination
- Access tokens typically expire after 1 hour (refresh as needed)

## Security Considerations

- **Token Storage**: Never commit access tokens to version control
- **Permissions**: Use minimum required permissions (Sites.Read.All vs
  Sites.ReadWrite.All)
- **Token Expiry**: Implement token refresh logic for long-running operations
- **Rate Limits**: Microsoft Graph API has rate limits; the agent respects these

## Troubleshooting

### Common Issues

1. **"Access token expired"**
   - Solution: Refresh your OAuth2 token before invoking the agent

2. **"Insufficient permissions"**
   - Solution: Ensure your app registration has Sites.Read.All or
     Sites.ReadWrite.All permissions
   - Verify admin consent has been granted

3. **"Site not found"**
   - Solution: Verify the site ID format is correct
   - Try using the site URL instead:
     `/sites/contoso.sharepoint.com:/sites/sitename`

4. **"Timeout"**
   - Solution: Reduce max_depth or increase maxTimeMinutes in configuration
   - Process smaller folder subsets

## Development

### Running Tests

```bash
cd packages/core
npm test -- sharepoint-graph.test.ts
```

### Building

```bash
npm run build
```

## Related Documentation

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/api/overview)
- [SharePoint API Reference](https://docs.microsoft.com/en-us/graph/api/resources/sharepoint)
- [OAuth 2.0 Authentication](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow)

## License

Copyright 2025 Google LLC. Licensed under the Apache License, Version 2.0.
