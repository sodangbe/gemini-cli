/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * Example script demonstrating how to use the SharePoint Folder Processor Agent
 *
 * This example shows how to:
 * 1. Configure the agent
 * 2. Obtain an access token (using MSAL)
 * 3. Invoke the agent to analyze a SharePoint folder structure
 * 4. Process the results
 */

// NOTE: This is a conceptual example. To run this, you would need:
// - A SharePoint site with appropriate permissions
// - Azure AD app registration with Sites.Read.All permissions
// - The @azure/msal-node package installed

import { Config } from '@google/gemini-cli-core';

/**
 * Example: Configure and use the SharePoint Folder Processor Agent
 */
async function analyzeSharePointFolders() {
  // 1. Configure Gemini CLI with SharePoint agent enabled
  const config = new Config({
    targetDir: process.cwd(),
    sessionId: 'example-session',
    sharepointFolderProcessorSettings: {
      enabled: true,
      maxNumTurns: 50,
      maxTimeMinutes: 10,
      thinkingBudget: -1,
      model: 'gemini-2.0-flash-exp',
    },
  });

  // 2. Obtain an access token (example using environment variables)
  // In production, use @azure/msal-node or similar OAuth library
  const accessToken = process.env.GRAPH_ACCESS_TOKEN;
  const siteId = process.env.SHAREPOINT_SITE_ID;

  if (!accessToken || !siteId) {
    console.error(
      'Please set GRAPH_ACCESS_TOKEN and SHAREPOINT_SITE_ID environment variables',
    );
    process.exit(1);
  }

  // 3. The agent is automatically available as a tool when enabled
  // You would typically invoke it through the Gemini client like this:

  console.log('SharePoint Folder Processor Agent Configuration:');
  console.log(
    '- Agent enabled:',
    config.getSharePointFolderProcessorSettings().enabled,
  );
  console.log(
    '- Max turns:',
    config.getSharePointFolderProcessorSettings().maxNumTurns,
  );
  console.log(
    '- Max time (minutes):',
    config.getSharePointFolderProcessorSettings().maxTimeMinutes,
  );
  console.log('\nTo invoke the agent, use it through the Gemini CLI:');
  console.log(`
    Example prompt:
    "Analyze the SharePoint folder structure at site ${siteId} 
     with access token ${accessToken.substring(0, 10)}...
     Start from the root folder and analyze up to 5 levels deep."
  `);

  // The agent will automatically:
  // - Connect to SharePoint using the Graph API
  // - Traverse the folder structure
  // - Collect metadata about folders and files
  // - Generate a comprehensive JSON report
}

/**
 * Example: Direct use of SharePoint Graph API Tool
 */
async function useGraphApiTool() {
  const config = new Config({
    targetDir: process.cwd(),
    sessionId: 'example-session',
  });

  const toolRegistry = await config.createToolRegistry();
  const graphTool = toolRegistry.getTool('sharepoint_graph_api');

  if (!graphTool) {
    console.error('SharePoint Graph API tool not found');
    return;
  }

  const accessToken = process.env.GRAPH_ACCESS_TOKEN;
  const siteId = process.env.SHAREPOINT_SITE_ID;

  if (!accessToken || !siteId) {
    console.error(
      'Please set GRAPH_ACCESS_TOKEN and SHAREPOINT_SITE_ID environment variables',
    );
    return;
  }

  // Example: List root folder items
  const invocation = graphTool.build({
    endpoint: `/sites/${siteId}/drive/root/children`,
    access_token: accessToken,
    method: 'GET',
  });

  const controller = new AbortController();
  const result = await invocation.execute(controller.signal);

  if (result.error) {
    console.error('Error:', result.error.message);
  } else {
    console.log('Success!');
    console.log('Result:', result.llmContent);
  }
}

/**
 * Example: Getting an access token with MSAL
 *
 * This function shows how to obtain an OAuth2 token for Microsoft Graph API
 * Note: Requires @azure/msal-node package
 */
async function getAccessTokenExample() {
  // This is a conceptual example - uncomment and install @azure/msal-node to use
  /*
  const msal = require('@azure/msal-node');

  const msalConfig = {
    auth: {
      clientId: process.env.AZURE_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
      clientSecret: process.env.AZURE_CLIENT_SECRET,
    },
  };

  const cca = new msal.ConfidentialClientApplication(msalConfig);

  const tokenRequest = {
    scopes: ['https://graph.microsoft.com/.default'],
  };

  try {
    const response = await cca.acquireTokenByClientCredential(tokenRequest);
    return response.accessToken;
  } catch (error) {
    console.error('Error acquiring token:', error);
    throw error;
  }
  */

  console.log(`
    To get an access token:
    
    1. Register an app in Azure AD:
       https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
    
    2. Add Microsoft Graph API permissions:
       - Sites.Read.All (for read-only access)
       - Sites.ReadWrite.All (for read-write access)
    
    3. Grant admin consent for the permissions
    
    4. Create a client secret
    
    5. Use the credentials with @azure/msal-node:
       npm install @azure/msal-node
    
    6. Set environment variables:
       export AZURE_CLIENT_ID="your-client-id"
       export AZURE_TENANT_ID="your-tenant-id"
       export AZURE_CLIENT_SECRET="your-client-secret"
  `);
}

// Main execution
if (require.main === module) {
  console.log('SharePoint Folder Processor Agent - Example Usage\n');
  console.log('='.repeat(60));

  getAccessTokenExample()
    .then(() => {
      console.log('\n' + '='.repeat(60));
      console.log(
        '\nFor actual usage, uncomment the appropriate example above',
      );
      console.log('and ensure you have valid credentials configured.\n');
    })
    .catch(console.error);
}

export { analyzeSharePointFolders, useGraphApiTool, getAccessTokenExample };
