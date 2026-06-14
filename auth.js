// auth.js — Microsoft 365 sign-in via MSAL (Authorization Code + PKCE, redirect flow).
// Requires the global `msal` (vendor/msal-browser.min.js) to be loaded first.
import { CONFIG } from "./config.js";

let pca = null;
let initPromise = null;

// Directory URL of the app (strip the filename) — must match a registered redirect URI.
const redirectUri = () => location.origin + location.pathname.replace(/[^/]*$/, "");

export function initAuth() {
  if (initPromise) return initPromise;
  initPromise = (async () => {
    pca = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: redirectUri(),
      },
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
    });
    await pca.initialize();
    const result = await pca.handleRedirectPromise(); // completes a returning sign-in, if any
    if (result && result.account) pca.setActiveAccount(result.account);
    else if (!pca.getActiveAccount()) {
      const accts = pca.getAllAccounts();
      if (accts.length) pca.setActiveAccount(accts[0]);
    }
    return getAccount();
  })();
  return initPromise;
}

export const getAccount = () => (pca ? pca.getActiveAccount() : null);

export async function signIn() {
  await pca.loginRedirect({ scopes: CONFIG.scopes }); // navigates away
}

export async function signOut() {
  await pca.logoutRedirect({ account: getAccount() });
}

// Returns a Graph access token, refreshing silently when possible.
export async function getToken() {
  const account = getAccount();
  if (!account) throw new Error("Not signed in");
  try {
    const r = await pca.acquireTokenSilent({ account, scopes: CONFIG.scopes });
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await pca.acquireTokenRedirect({ scopes: CONFIG.scopes }); // navigates away
    }
    throw e;
  }
}
