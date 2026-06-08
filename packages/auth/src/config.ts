/**
 * Auth configuration resolution for the Microsoft MCP servers.
 *
 * The servers authenticate to Microsoft Graph with MSAL (a public-client OAuth
 * flow). The end user registers their own Entra ID app and supplies its client
 * id via {@link CLIENT_ID_ENV}; tokens are cached on disk and refreshed
 * silently. This module is pure (env in, config out) so it is fully unit-tested;
 * the MSAL/network/browser interaction lives in `./msal.ts`.
 */
import { homedir } from "node:os";
import { join } from "node:path";

/** Env var holding the Entra ID application (client) id. Required for sign-in. */
export const CLIENT_ID_ENV = "MICROSOFT_CLIENT_ID";
/** Env var selecting the tenant (`common`, `organizations`, `consumers`, or a tenant id). */
export const TENANT_ENV = "MICROSOFT_TENANT_ID";
/** Env var overriding the directory used for the token cache. */
export const CACHE_DIR_ENV = "MICROSOFT_MCP_CACHE_DIR";

export interface AuthConfig {
  clientId: string;
  /** Full authority URL, e.g. `https://login.microsoftonline.com/common`. */
  authority: string;
  /** Absolute path to the on-disk MSAL token cache. */
  cacheFile: string;
}

/** Thrown when {@link CLIENT_ID_ENV} is not set. Carries a 401 status. */
export class AuthConfigError extends Error {
  readonly status = 401;
  constructor(message: string) {
    super(message);
    this.name = "AuthConfigError";
  }
}

/** Thrown when there is no cached sign-in — the user must run `<bin> login`. 401. */
export class NotSignedInError extends Error {
  readonly status = 401;
  constructor(message: string) {
    super(message);
    this.name = "NotSignedInError";
  }
}

/** Directory holding the token cache: `$MICROSOFT_MCP_CACHE_DIR`, else XDG, else `~/.config/microsoft-mcp`. */
export const configDir = (env: NodeJS.ProcessEnv = process.env): string => {
  if (env[CACHE_DIR_ENV]) return env[CACHE_DIR_ENV] as string;
  if (env.XDG_CONFIG_HOME) return join(env.XDG_CONFIG_HOME, "microsoft-mcp");
  return join(homedir(), ".config", "microsoft-mcp");
};

/** Absolute path to the MSAL token cache file. */
export const cacheFilePath = (env: NodeJS.ProcessEnv = process.env): string =>
  join(configDir(env), "token-cache.json");

/** Builds the authority URL for a tenant (defaults to `common`). */
export const authorityFor = (tenant: string | undefined): string =>
  `https://login.microsoftonline.com/${tenant && tenant.trim() ? tenant.trim() : "common"}`;

/**
 * Resolves {@link AuthConfig} from the environment. Throws {@link AuthConfigError}
 * if no client id is configured.
 */
export const resolveAuthConfig = (env: NodeJS.ProcessEnv = process.env): AuthConfig => {
  const clientId = env[CLIENT_ID_ENV];
  if (!clientId) {
    throw new AuthConfigError(
      `Microsoft sign-in is not configured. Set ${CLIENT_ID_ENV} to your Entra ID app's ` +
        `Application (client) ID, then run the server's \`login\` command. ` +
        `See the README "Authentication" section.`,
    );
  }
  return {
    clientId,
    authority: authorityFor(env[TENANT_ENV]),
    cacheFile: cacheFilePath(env),
  };
};
