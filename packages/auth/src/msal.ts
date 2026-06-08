/**
 * MSAL (public-client) interaction: silent token acquisition from the on-disk
 * cache, plus interactive sign-in via the browser (auth code + PKCE on a
 * loopback redirect) with a device-code fallback.
 *
 * Excluded from coverage — every path here is network / browser / OS interaction
 * that can't be exercised in a unit test. The pure config logic lives in
 * `./config.ts` and is tested there.
 */
import { createServer, type Server } from "node:http";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import { exec } from "node:child_process";
import {
  PublicClientApplication,
  CryptoProvider,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import { log } from "@microsoft-mcp/logger";
import { type AuthConfig, NotSignedInError } from "./config.js";

const cachePlugin = (cacheFile: string): ICachePlugin => ({
  beforeCacheAccess: async (ctx: TokenCacheContext) => {
    try {
      ctx.tokenCache.deserialize(await readFile(cacheFile, "utf8"));
    } catch {
      // No cache yet — first run.
    }
  },
  afterCacheAccess: async (ctx: TokenCacheContext) => {
    if (ctx.cacheHasChanged) {
      await mkdir(dirname(cacheFile), { recursive: true });
      await writeFile(cacheFile, ctx.tokenCache.serialize(), { mode: 0o600 });
    }
  },
});

const createPca = (config: AuthConfig): PublicClientApplication => {
  const msalConfig: Configuration = {
    auth: { clientId: config.clientId, authority: config.authority },
    cache: { cachePlugin: cachePlugin(config.cacheFile) },
  };
  return new PublicClientApplication(msalConfig);
};

/** Acquire an access token silently, refreshing from the cached refresh token. */
export const acquireTokenSilent = async (config: AuthConfig, scopes: string[]): Promise<string> => {
  const pca = createPca(config);
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0];
  if (!account) {
    throw new NotSignedInError(
      "Not signed in to Microsoft. Run this server's `login` command once to authenticate " +
        "(e.g. `npx <package> login`), then start the server again.",
    );
  }
  let result;
  try {
    result = await pca.acquireTokenSilent({ account, scopes });
  } catch (err: any) {
    throw new NotSignedInError(
      `Could not refresh your Microsoft sign-in (${err?.message ?? "unknown error"}). ` +
        "Run the `login` command again.",
    );
  }
  if (!result?.accessToken) {
    throw new NotSignedInError("No access token returned; run the `login` command again.");
  }
  return result.accessToken;
};

const openBrowser = (url: string): void => {
  const cmd =
    process.platform === "darwin"
      ? "open"
      : process.platform === "win32"
        ? "start \"\""
        : "xdg-open";
  exec(`${cmd} "${url}"`, () => {
    // best-effort; the URL is also printed for manual opening
  });
};

/** Interactive browser sign-in: auth code + PKCE on an ephemeral loopback port. */
const loginViaBrowser = async (config: AuthConfig, scopes: string[]): Promise<void> => {
  const pca = createPca(config);
  const crypto = new CryptoProvider();
  const { verifier, challenge } = await crypto.generatePkceCodes();

  await new Promise<void>((resolve, reject) => {
    let redirectUri = "";
    const server: Server = createServer((req, res) => {
      const url = new URL(req.url ?? "/", redirectUri || "http://127.0.0.1");
      const error = url.searchParams.get("error");
      const code = url.searchParams.get("code");
      if (error) {
        res.writeHead(400, { "Content-Type": "text/html" }).end(`<h2>Sign-in failed: ${error}</h2>`);
        server.close();
        reject(new Error(`${error}: ${url.searchParams.get("error_description") ?? ""}`));
        return;
      }
      if (!code) {
        res.writeHead(404).end();
        return;
      }
      pca
        .acquireTokenByCode({ code, scopes, redirectUri, codeVerifier: verifier })
        .then(() => {
          res
            .writeHead(200, { "Content-Type": "text/html" })
            .end("<h2>Signed in ✓</h2><p>You can close this tab and return to your terminal.</p>");
          server.close();
          resolve();
        })
        .catch((err) => {
          res.writeHead(500, { "Content-Type": "text/html" }).end(`<h2>Sign-in failed</h2><pre>${err?.message}</pre>`);
          server.close();
          reject(err);
        });
    });

    server.on("error", reject);
    server.listen(0, "127.0.0.1", async () => {
      try {
        const { port } = server.address() as { port: number };
        redirectUri = `http://localhost:${port}`;
        const authUrl = await pca.getAuthCodeUrl({
          scopes,
          redirectUri,
          codeChallenge: challenge,
          codeChallengeMethod: "S256",
        });
        console.log(`\nOpening your browser to sign in to Microsoft...\nIf it doesn't open, paste this URL:\n\n  ${authUrl}\n`);
        openBrowser(authUrl);
      } catch (err) {
        server.close();
        reject(err);
      }
    });
  });
};

/** Device-code sign-in: prints a code to enter at microsoft.com/devicelogin. */
const loginViaDeviceCode = async (config: AuthConfig, scopes: string[]): Promise<void> => {
  const pca = createPca(config);
  await pca.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (resp) => {
      console.log(`\n${resp.message}\n`);
    },
  });
};

/**
 * Interactive sign-in. Defaults to the browser flow and falls back to device
 * code if the browser/loopback flow fails (or when {@link useDeviceCode} is set).
 */
export const login = async (config: AuthConfig, scopes: string[], useDeviceCode = false): Promise<void> => {
  if (useDeviceCode) {
    await loginViaDeviceCode(config, scopes);
    return;
  }
  try {
    await loginViaBrowser(config, scopes);
  } catch (err: any) {
    log.warn("browser sign-in failed; falling back to device code", { message: err?.message });
    await loginViaDeviceCode(config, scopes);
  }
};
