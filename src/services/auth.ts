import {
  PublicClientApplication,
  DeviceCodeRequest,
  AuthenticationResult,
  AccountInfo,
  Configuration,
  TokenCacheContext,
} from "@azure/msal-node";
import { SCOPES } from "../constants.js";
import * as fs from "fs";
import * as path from "path";

const TOKEN_CACHE_PATH = path.join(
  process.env.HOME || process.env.USERPROFILE || ".",
  ".msgraph-mcp-token-cache.json"
);

let msalInstance: PublicClientApplication | null = null;
let cachedAccount: AccountInfo | null = null;

function getMsalConfig(): Configuration {
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const tenantId = process.env.MICROSOFT_TENANT_ID || "common";

  if (!clientId) {
    throw new Error(
      "MICROSOFT_CLIENT_ID environment variable is required. " +
        "Register an app in Azure Portal (Entra ID > App registrations) " +
        "and set MICROSOFT_CLIENT_ID to its Application (client) ID. " +
        "Ensure 'Allow public client flows' is enabled."
    );
  }

  return {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (cacheContext: TokenCacheContext) => {
          if (fs.existsSync(TOKEN_CACHE_PATH)) {
            const data = fs.readFileSync(TOKEN_CACHE_PATH, "utf-8");
            cacheContext.tokenCache.deserialize(data);
          }
        },
        afterCacheAccess: async (cacheContext: TokenCacheContext) => {
          if (cacheContext.cacheHasChanged) {
            fs.writeFileSync(
              TOKEN_CACHE_PATH,
              cacheContext.tokenCache.serialize(),
              "utf-8"
            );
          }
        },
      },
    },
  };
}

async function getMsalInstance(): Promise<PublicClientApplication> {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication(getMsalConfig());
  }
  return msalInstance;
}

async function getCachedAccount(): Promise<AccountInfo | null> {
  if (cachedAccount) return cachedAccount;
  const pca = await getMsalInstance();
  const cache = pca.getTokenCache();
  const accounts = await cache.getAllAccounts();
  if (accounts.length > 0) {
    cachedAccount = accounts[0];
    return cachedAccount;
  }
  return null;
}

/**
 * Get an access token, using silent acquisition if possible,
 * falling back to device code flow if not.
 */
export async function getAccessToken(): Promise<string> {
  const pca = await getMsalInstance();
  const account = await getCachedAccount();

  // Try silent token acquisition first
  if (account) {
    try {
      const result: AuthenticationResult = await pca.acquireTokenSilent({
        account,
        scopes: SCOPES.filter((s) => s !== "offline_access"),
      });
      return result.accessToken;
    } catch {
      // Silent acquisition failed, fall through to device code
      console.error(
        "[auth] Silent token acquisition failed, initiating device code flow..."
      );
    }
  }

  // Device code flow
  const deviceCodeRequest: DeviceCodeRequest = {
    scopes: SCOPES.filter((s) => s !== "offline_access"),
    deviceCodeCallback: (response) => {
      // Print to stderr so it doesn't interfere with MCP stdio transport
      console.error("\n" + "=".repeat(60));
      console.error("🔐 Microsoft Authentication Required");
      console.error("=".repeat(60));
      console.error(response.message);
      console.error("=".repeat(60) + "\n");
    },
  };

  const result = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
  if (!result) {
    throw new Error("Device code authentication failed - no result returned");
  }
  if (result.account) {
    cachedAccount = result.account;
  }
  return result.accessToken;
}

/**
 * Clear the token cache and force re-authentication
 */
export async function clearTokenCache(): Promise<void> {
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    fs.unlinkSync(TOKEN_CACHE_PATH);
  }
  cachedAccount = null;
  msalInstance = null;
}

/**
 * Check if we have a cached account (i.e., previously authenticated)
 */
export async function isAuthenticated(): Promise<boolean> {
  const account = await getCachedAccount();
  return account !== null;
}
