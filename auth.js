/**
 * auth.js
 * -------
 * MSAL authentication module.
 * Handles sign-in, sign-out, and token acquisition.
 * Uses redirect flow (same pattern as colleague's working dashboard).
 *
 * Token is passed to the backend API in the Authorization header.
 * The backend uses it to call Microsoft Graph on behalf of the user.
 */

'use strict';

const AUTH_CONFIG = {
  clientId:    '7dfd41b7-1960-49c6-a701-b021a28550bb', // WFM Reporting Data app
  tenantId:    '3acfc3b8-ea64-4ce1-8910-f91e9c67a3fc',
  redirectUri: window.location.origin + window.location.pathname,
};

const SCOPES = ['User.Read', 'Files.ReadWrite.All', 'Sites.ReadWrite.All'];

let _msalInstance = null;
let _cachedToken  = null;

// ── MSAL instance (lazy, same pattern as colleague) ───────────────────────────
function getMsal() {
  if (_msalInstance) return _msalInstance;
  _msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    AUTH_CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${AUTH_CONFIG.tenantId}`,
      redirectUri: AUTH_CONFIG.redirectUri,
    },
    cache: { cacheLocation: 'sessionStorage' },
  });
  return _msalInstance;
}

// ── Public: Sign in (redirect) ────────────────────────────────────────────────
async function signIn() {
  try {
    const i = getMsal();
    await i.initialize();
    await i.loginRedirect({ scopes: SCOPES });
  } catch (e) {
    alert('Sign-in failed: ' + e.message);
  }
}

// ── Public: Sign out ──────────────────────────────────────────────────────────
function signOut() {
  getMsal().logoutRedirect({ postLogoutRedirectUri: AUTH_CONFIG.redirectUri });
}

// ── Public: Get access token (silent, falls back to cache) ────────────────────
async function getToken() {
  const i = getMsal();
  const account = i.getActiveAccount() || i.getAllAccounts()[0];
  if (!account) throw new Error('Not signed in');

  try {
    const result = await i.acquireTokenSilent({ scopes: SCOPES, account });
    _cachedToken = result.accessToken;
    return _cachedToken;
  } catch (e) {
    if (_cachedToken) return _cachedToken; // use cache if silent fails
    await i.acquireTokenRedirect({ scopes: SCOPES, account });
    return null;
  }
}

// ── Public: Get current user info ─────────────────────────────────────────────
function getCurrentUser() {
  const i = getMsal();
  return i.getActiveAccount() || i.getAllAccounts()[0] || null;
}

// ── Bootstrap: handle redirect on page load ───────────────────────────────────
(async function initAuth() {
  try {
    const i = getMsal();
    await i.initialize();

    // Handle redirect response (coming back from Microsoft login)
    const result = await i.handleRedirectPromise();
    if (result?.account) {
      i.setActiveAccount(result.account);
      onAuthSuccess(result.account);
      return;
    }

    // Check existing session
    const accounts = i.getAllAccounts();
    if (accounts.length > 0) {
      i.setActiveAccount(accounts[0]);
      onAuthSuccess(accounts[0]);
    }
    // else: stay on login page
  } catch (e) {
    console.error('Auth init error:', e);
  }
})();

// ── Called after successful auth ──────────────────────────────────────────────
function onAuthSuccess(account) {
  // Update UI
  document.getElementById('hdr-user').textContent = account.name || account.username;

  // Show app, hide login
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-app').classList.add('active');

  // Bootstrap the dashboard (defined in app.js)
  if (typeof onSignedIn === 'function') onSignedIn(account);
}

// Expose to HTML onclick handlers
window.signIn  = signIn;
window.signOut = signOut;
