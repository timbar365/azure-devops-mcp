// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as https from "node:https";
import type { INetworkModule, NetworkRequestOptions, NetworkResponse } from "@azure/msal-common/node";
import { HttpsProxyAgent } from "https-proxy-agent";
import type { IProxyConfiguration } from "typed-rest-client/Interfaces.js";
import { logger } from "./logger.js";

/**
 * Returns the proxy URL from standard environment variables.
 * Checks HTTPS_PROXY, https_proxy, HTTP_PROXY, http_proxy in that order.
 */
export function getProxyUrl(): string | undefined {
  return process.env["HTTPS_PROXY"] ?? process.env["https_proxy"] ?? process.env["HTTP_PROXY"] ?? process.env["http_proxy"];
}

/**
 * Returns the list of hosts that should bypass the proxy,
 * parsed from NO_PROXY / no_proxy environment variables.
 */
function getNoProxyHosts(): string[] {
  const noProxy = process.env["NO_PROXY"] ?? process.env["no_proxy"];
  if (!noProxy) return [];
  return noProxy
    .split(",")
    .map((h) => h.trim())
    .filter(Boolean);
}

/**
 * Builds a typed-rest-client IProxyConfiguration from the environment,
 * or returns undefined if no proxy is configured.
 */
export function getProxyConfig(): IProxyConfiguration | undefined {
  const proxyUrl = getProxyUrl();
  if (!proxyUrl) return undefined;

  try {
    const parsed = new URL(proxyUrl);
    const config: IProxyConfiguration = {
      proxyUrl: `${parsed.protocol}//${parsed.host}`,
    };
    if (parsed.username) config.proxyUsername = decodeURIComponent(parsed.username);
    if (parsed.password) config.proxyPassword = decodeURIComponent(parsed.password);

    const bypass = getNoProxyHosts();
    if (bypass.length > 0) config.proxyBypassHosts = bypass;

    logger.debug(`Proxy configured for azure-devops-node-api: ${config.proxyUrl}`);
    return config;
  } catch {
    logger.warn(`Invalid proxy URL in environment: ${proxyUrl}`);
    return undefined;
  }
}

/**
 * Makes an HTTPS request routing through the configured proxy agent when
 * HTTPS_PROXY / HTTP_PROXY is set, otherwise delegates to the global fetch.
 */
export async function proxyFetch(url: string, method: string, headers: Record<string, string> = {}, body?: string): Promise<{ status: number; headers: Record<string, string>; body: string }> {
  const proxyUrl = getProxyUrl();

  if (!proxyUrl) {
    // No proxy configured — use the built-in fetch so tests can mock it easily.
    const response = await fetch(url, { method, headers, body });
    const resHeaders: Record<string, string> = {};
    response.headers.forEach((value, key) => {
      resHeaders[key] = value;
    });
    const responseBody = await response.text();
    return { status: response.status, headers: resHeaders, body: responseBody };
  }

  // Proxy configured — use https-proxy-agent with node:https directly.
  logger.debug(`proxyFetch: routing ${method} ${url} through proxy`);
  const agent = new HttpsProxyAgent(proxyUrl);

  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const options: https.RequestOptions = {
      hostname: parsed.hostname,
      port: parsed.port || 443,
      path: parsed.pathname + parsed.search,
      method,
      headers,
      agent,
    };

    const req = https.request(options, (res) => {
      const resHeaders: Record<string, string> = {};
      for (const [key, value] of Object.entries(res.headers)) {
        if (typeof value === "string") resHeaders[key] = value;
        else if (Array.isArray(value)) resHeaders[key] = value.join(", ");
      }

      let data = "";
      res.on("data", (chunk: Buffer) => (data += chunk.toString()));
      res.on("end", () => resolve({ status: res.statusCode ?? 0, headers: resHeaders, body: data }));
    });

    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

/**
 * Custom MSAL INetworkModule that routes requests through the configured proxy.
 * Used when HTTPS_PROXY / HTTP_PROXY is set in the environment.
 */
class ProxyNetworkClient implements INetworkModule {
  async sendGetRequestAsync<T>(url: string, options?: NetworkRequestOptions, _timeout?: number): Promise<NetworkResponse<T>> {
    const headers = (options?.headers as Record<string, string> | undefined) ?? {};
    const result = await proxyFetch(url, "GET", headers);
    return {
      headers: result.headers,
      body: JSON.parse(result.body) as T,
      status: result.status,
    };
  }

  async sendPostRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
    const headers = (options?.headers as Record<string, string> | undefined) ?? {};
    const result = await proxyFetch(url, "POST", headers, options?.body);
    return {
      headers: result.headers,
      body: JSON.parse(result.body) as T,
      status: result.status,
    };
  }
}

/**
 * Returns a proxy-aware MSAL network client when a proxy URL is configured,
 * or undefined to use MSAL's built-in fetch-based client.
 */
export function getMsalNetworkClient(): INetworkModule | undefined {
  const proxyUrl = getProxyUrl();
  if (!proxyUrl) return undefined;

  logger.debug(`Proxy configured for MSAL network client: ${proxyUrl}`);
  return new ProxyNetworkClient();
}
