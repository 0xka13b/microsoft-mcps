export {
  resolveAuthConfig,
  configDir,
  cacheFilePath,
  authorityFor,
  AuthConfigError,
  NotSignedInError,
  CLIENT_ID_ENV,
  TENANT_ENV,
  CACHE_DIR_ENV,
} from "./config.js";
export type { AuthConfig } from "./config.js";
export { acquireTokenSilent, login } from "./msal.js";

import { resolveAuthConfig } from "./config.js";
import { acquireTokenSilent } from "./msal.js";

/**
 * High-level helper: resolve config from the environment and return an access
 * token for the given scopes, refreshing silently from the on-disk cache.
 * Throws {@link AuthConfigError} / {@link NotSignedInError} (both status 401).
 */
export const acquireToken = (scopes: string[]): Promise<string> =>
  acquireTokenSilent(resolveAuthConfig(), scopes);
