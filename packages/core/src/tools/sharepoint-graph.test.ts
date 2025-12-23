/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { SharePointGraphTool } from './sharepoint-graph.js';
import type { Config } from '../config/config.js';
import { ApprovalMode } from '../config/config.js';
import { ToolErrorType } from './tool-error.js';

const mockConfig = {
  getApprovalMode: vi.fn(() => ApprovalMode.DEFAULT),
  getDebugMode: vi.fn(() => false),
  getEnableMessageBusIntegration: vi.fn(() => false),
} as unknown as Config;

// Mock global fetch
global.fetch = vi.fn();

describe('SharePointGraphTool', () => {
  let tool: SharePointGraphTool;

  beforeEach(() => {
    vi.clearAllMocks();
    tool = new SharePointGraphTool(mockConfig);
  });

  describe('validateToolParamValues', () => {
    it('should reject empty endpoint', () => {
      const params = {
        endpoint: '',
        access_token: 'test-token',
      };

      const invocation = () => tool.build(params);
      expect(invocation).toThrow("The 'endpoint' parameter must be non-empty");
    });

    it('should reject empty access_token', () => {
      const params = {
        endpoint: '/sites/test',
        access_token: '',
      };

      const invocation = () => tool.build(params);
      expect(invocation).toThrow(
        "The 'access_token' parameter must be non-empty",
      );
    });

    it('should reject endpoint without leading slash', () => {
      const params = {
        endpoint: 'sites/test',
        access_token: 'test-token',
      };

      const invocation = () => tool.build(params);
      expect(invocation).toThrow("The 'endpoint' parameter must start with");
    });

    it('should reject invalid HTTP method', () => {
      const params = {
        endpoint: '/sites/test',
        access_token: 'test-token',
        method: 'INVALID',
      };

      const invocation = () => tool.build(params);
      expect(invocation).toThrow('Invalid HTTP method');
    });

    it('should reject invalid JSON body', () => {
      const params = {
        endpoint: '/sites/test',
        access_token: 'test-token',
        method: 'POST',
        body: 'not valid json',
      };

      const invocation = () => tool.build(params);
      expect(invocation).toThrow("The 'body' parameter must be valid JSON");
    });

    it('should accept valid parameters', () => {
      const params = {
        endpoint: '/sites/test/drive/root/children',
        access_token: 'test-token',
      };

      const invocation = tool.build(params);
      expect(invocation).toBeDefined();
      expect(invocation.params).toEqual(params);
    });

    it('should accept valid parameters with method and body', () => {
      const params = {
        endpoint: '/sites/test/drive/items',
        access_token: 'test-token',
        method: 'POST',
        body: JSON.stringify({ name: 'test' }),
      };

      const invocation = tool.build(params);
      expect(invocation).toBeDefined();
      expect(invocation.params).toEqual(params);
    });
  });

  describe('execute', () => {
    it('should make successful Graph API call', async () => {
      const mockResponse = {
        value: [
          {
            id: '123',
            name: 'Test Folder',
            folder: { childCount: 5 },
          },
        ],
      };

      (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
        ok: true,
        json: async () => mockResponse,
      });

      const params = {
        endpoint: '/sites/test/drive/root/children',
        access_token: 'test-token',
      };

      const invocation = tool.build(params);
      const result = await invocation.execute(new AbortController().signal);

      expect(result.error).toBeUndefined();
      expect(result.llmContent).toContain('Test Folder');
      expect(result.returnDisplay).toContain('Successfully retrieved data');
      expect(global.fetch).toHaveBeenCalledWith(
        expect.stringContaining('/sites/test/drive/root/children'),
        expect.objectContaining({
          method: 'GET',
          headers: expect.objectContaining({
            Authorization: 'Bearer test-token',
          }),
        }),
      );
    });

    it('should handle HTTP error responses', async () => {
      const mockErrorResponse = {
        error: {
          message: 'Resource not found',
        },
      };

      (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
        ok: false,
        status: 404,
        statusText: 'Not Found',
        json: async () => mockErrorResponse,
      });

      const params = {
        endpoint: '/sites/invalid/drive/root',
        access_token: 'test-token',
      };

      const invocation = tool.build(params);
      const result = await invocation.execute(new AbortController().signal);

      expect(result.error).toBeDefined();
      expect(result.error?.type).toBe(ToolErrorType.WEB_FETCH_PROCESSING_ERROR);
      expect(result.llmContent).toContain('Resource not found');
    });

    it('should handle network errors', async () => {
      (global.fetch as ReturnType<typeof vi.fn>).mockRejectedValueOnce(
        new Error('Network error'),
      );

      const params = {
        endpoint: '/sites/test/drive/root/children',
        access_token: 'test-token',
      };

      const invocation = tool.build(params);
      const result = await invocation.execute(new AbortController().signal);

      expect(result.error).toBeDefined();
      expect(result.error?.type).toBe(ToolErrorType.WEB_FETCH_PROCESSING_ERROR);
      expect(result.llmContent).toContain('Network error');
    });

    it('should support POST requests with body', async () => {
      const mockResponse = {
        id: '456',
        name: 'Created Item',
      };

      (global.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
        ok: true,
        json: async () => mockResponse,
      });

      const params = {
        endpoint: '/sites/test/drive/items',
        access_token: 'test-token',
        method: 'POST',
        body: JSON.stringify({ name: 'New Item' }),
      };

      const invocation = tool.build(params);
      const result = await invocation.execute(new AbortController().signal);

      expect(result.error).toBeUndefined();
      expect(result.llmContent).toContain('Created Item');
      expect(global.fetch).toHaveBeenCalledWith(
        expect.stringContaining('/sites/test/drive/items'),
        expect.objectContaining({
          method: 'POST',
          body: params.body,
        }),
      );
    });
  });

  describe('getDescription', () => {
    it('should return description with endpoint', () => {
      const params = {
        endpoint: '/sites/test/drive/root/children',
        access_token: 'test-token',
      };

      const invocation = tool.build(params);
      const description = invocation.getDescription();

      expect(description).toContain('GET');
      expect(description).toContain('/sites/test/drive/root/children');
    });

    it('should include custom method in description', () => {
      const params = {
        endpoint: '/sites/test/drive/items',
        access_token: 'test-token',
        method: 'POST',
      };

      const invocation = tool.build(params);
      const description = invocation.getDescription();

      expect(description).toContain('POST');
      expect(description).toContain('/sites/test/drive/items');
    });
  });
});
