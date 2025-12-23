/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import type {
  ToolCallConfirmationDetails,
  ToolInvocation,
  ToolResult,
} from './tools.js';
import { BaseDeclarativeTool, BaseToolInvocation, Kind } from './tools.js';
import type { MessageBus } from '../confirmation-bus/message-bus.js';
import { ToolErrorType } from './tool-error.js';
import { getErrorMessage } from '../utils/errors.js';
import type { Config } from '../config/config.js';
import { ApprovalMode } from '../config/config.js';
import { SHAREPOINT_GRAPH_TOOL_NAME } from './tool-names.js';

const GRAPH_API_TIMEOUT_MS = 30000;
const GRAPH_API_BASE_URL = 'https://graph.microsoft.com/v1.0';

/**
 * Parameters for the SharePoint Graph API tool
 */
export interface SharePointGraphToolParams {
  /**
   * The Microsoft Graph API endpoint path (e.g., '/sites/{site-id}/drive/root/children')
   */
  endpoint: string;

  /**
   * The access token for Microsoft Graph API authentication
   */
  access_token: string;

  /**
   * Optional: HTTP method (GET, POST, etc.)
   */
  method?: string;

  /**
   * Optional: Request body for POST/PATCH requests
   */
  body?: string;
}

interface DriveItem {
  id: string;
  name: string;
  size?: number;
  folder?: {
    childCount: number;
  };
  file?: {
    mimeType: string;
  };
  webUrl: string;
  parentReference?: {
    path: string;
  };
}

interface GraphApiResponse {
  value?: DriveItem[];
  '@odata.nextLink'?: string;
  error?: {
    code: string;
    message: string;
  };
}

class SharePointGraphToolInvocation extends BaseToolInvocation<
  SharePointGraphToolParams,
  ToolResult
> {
  constructor(
    private config: Config,
    params: SharePointGraphToolParams,
    messageBus?: MessageBus,
  ) {
    super(params, messageBus);
  }

  getDescription(): string {
    return `Calling Microsoft Graph API: ${this.params.method || 'GET'} ${this.params.endpoint}`;
  }

  override async shouldConfirmExecute(
    abortSignal: AbortSignal,
  ): Promise<ToolCallConfirmationDetails | false> {
    // Try message bus confirmation first if available
    if (this.messageBus) {
      const decision = await this.getMessageBusDecision(abortSignal);
      if (decision === 'ALLOW') {
        return false; // No confirmation needed
      }
      if (decision === 'DENY') {
        throw new Error('SharePoint Graph API call denied by policy.');
      }
      // if 'ASK_USER', fall through to legacy logic
    }

    // Legacy confirmation flow (no message bus OR policy decision was ASK_USER)
    const approvalMode = this.config.getApprovalMode();
    if (approvalMode === ApprovalMode.AUTO_EDIT) {
      return false;
    }

    return false;
  }

  async execute(signal: AbortSignal): Promise<ToolResult> {
    try {
      const url = `${GRAPH_API_BASE_URL}${this.params.endpoint}`;
      const method = this.params.method?.toUpperCase() || 'GET';

      const headers: Record<string, string> = {
        Authorization: `Bearer ${this.params.access_token}`,
        'Content-Type': 'application/json',
      };

      const controller = new AbortController();
      const timeoutId = setTimeout(
        () => controller.abort(),
        GRAPH_API_TIMEOUT_MS,
      );

      // Combine the timeout signal with the provided abort signal
      const combinedSignal = signal.aborted
        ? signal
        : AbortSignal.any
          ? AbortSignal.any([signal, controller.signal])
          : controller.signal;

      const options: RequestInit = {
        method,
        headers,
        signal: combinedSignal,
      };

      if (this.params.body && (method === 'POST' || method === 'PATCH')) {
        options.body = this.params.body;
      }

      try {
        const response = await fetch(url, options);
        clearTimeout(timeoutId);

        if (!response.ok) {
          const errorData = (await response.json().catch(() => ({}))) as {
            error?: { message?: string };
          };
          const errorMessage =
            errorData.error?.message ||
            `HTTP ${response.status}: ${response.statusText}`;

          return {
            llmContent: `Error calling Microsoft Graph API: ${errorMessage}`,
            returnDisplay: `Error: ${errorMessage}`,
            error: {
              message: errorMessage,
              type: ToolErrorType.WEB_FETCH_PROCESSING_ERROR,
            },
          };
        }

        const data = (await response.json()) as GraphApiResponse;

        // Format the response for better readability
        const formattedContent = JSON.stringify(data, null, 2);

        return {
          llmContent: formattedContent,
          returnDisplay: `Successfully retrieved data from Microsoft Graph API\n${formattedContent}`,
        };
      } finally {
        clearTimeout(timeoutId);
      }
    } catch (error) {
      const errorMessage = getErrorMessage(error);
      return {
        llmContent: `Error calling Microsoft Graph API: ${errorMessage}`,
        returnDisplay: `Error: ${errorMessage}`,
        error: {
          message: errorMessage,
          type: ToolErrorType.WEB_FETCH_PROCESSING_ERROR,
        },
      };
    }
  }
}

/**
 * Tool for making Microsoft Graph API calls to access SharePoint data
 */
export class SharePointGraphTool extends BaseDeclarativeTool<
  SharePointGraphToolParams,
  ToolResult
> {
  constructor(
    private config: Config,
    messageBus?: MessageBus,
  ) {
    super(
      SHAREPOINT_GRAPH_TOOL_NAME,
      'SharePointGraph',
      `Makes authenticated calls to the Microsoft Graph API to access SharePoint data. Use this tool to retrieve information about SharePoint sites, document libraries, folders, and files. Supports various endpoints for navigating nested folder structures.`,
      Kind.Read,
      {
        properties: {
          endpoint: {
            description:
              "The Microsoft Graph API endpoint path (e.g., '/sites/{site-id}/drive/root/children' to list root items, '/sites/{site-id}/drive/items/{item-id}/children' to list children of a specific folder)",
            type: 'string',
          },
          access_token: {
            description:
              'The OAuth2 access token for Microsoft Graph API authentication. Must have appropriate permissions (Sites.Read.All or Sites.ReadWrite.All)',
            type: 'string',
          },
          method: {
            description:
              "Optional: HTTP method to use (default: 'GET'). Supported values: GET, POST, PATCH, DELETE",
            type: 'string',
          },
          body: {
            description:
              'Optional: JSON string containing the request body for POST/PATCH requests',
            type: 'string',
          },
        },
        required: ['endpoint', 'access_token'],
        type: 'object',
      },
      true,
      false,
      messageBus,
    );
  }

  protected override validateToolParamValues(
    params: SharePointGraphToolParams,
  ): string | null {
    if (!params.endpoint || params.endpoint.trim() === '') {
      return "The 'endpoint' parameter must be non-empty.";
    }

    if (!params.access_token || params.access_token.trim() === '') {
      return "The 'access_token' parameter must be non-empty.";
    }

    // Validate endpoint starts with '/'
    if (!params.endpoint.startsWith('/')) {
      return "The 'endpoint' parameter must start with '/'.";
    }

    // Validate method if provided
    if (params.method) {
      const validMethods = ['GET', 'POST', 'PATCH', 'DELETE'];
      if (!validMethods.includes(params.method.toUpperCase())) {
        return `Invalid HTTP method. Supported methods: ${validMethods.join(', ')}`;
      }
    }

    // Validate body is valid JSON if provided
    if (params.body) {
      try {
        JSON.parse(params.body);
      } catch (_e) {
        return "The 'body' parameter must be valid JSON.";
      }
    }

    return null;
  }

  protected createInvocation(
    params: SharePointGraphToolParams,
    messageBus?: MessageBus,
  ): ToolInvocation<SharePointGraphToolParams, ToolResult> {
    return new SharePointGraphToolInvocation(this.config, params, messageBus);
  }
}
