/**
 * auth.js
 * -------
 * MSAL v2 authentication for GitHub Pages SPA.
 * Always authenticates via Azure AD — the user's token is used to call
 * Microsoft Graph directly from the browser to read SharePoint Excel data.
 */

'use strict';

// ── Azure AD app registration ────────────────────────────────────────────────
const AUTH_CONFIG = {
  clientId:  '7dfd41b7-1960-49c6-a701-b021a28550bb',
  tenantId:  '3acfc3b8-ea64-4ce1-8910-f91e9c67a3fc',
  // Dynamically build redirect URI from current page URL
  get redirectUri() {
    return window.location.origin + window.location.pathname;
  },
};

// Scopes needed to read SharePoint Excel files via Graph API
const GRAPH_SCOPES = [
  'User.Read',
  'Files.Read.All',
  'Sites.Read.All',
];

let _msalInstance = null;
let _currentAccount = null;

// ── Initialize MSAL (v2 API) ────────────────────────────────────────────────
function getMsal() {
  if (_msalInstance) return _msalInstance;
  _msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    AUTH_CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${AUTH_CONFIG.tenantId}`,
      redirectUri: AUTH_CONFIG.redirectUri,
    },
    cache: {
      cacheLocation: 'sessionStorage',
      storeAuthStateInCookie: false,
    },
  });
  return _msalInstance;
}

// ── Get access token (silent first, then redirect) ──────────────────────────
window.getToken = async function () {
  const instance = getMsal();
  const account = _currentAccount || instance.getActiveAccount() || instance.getAllAccounts()[0];
  if (!account) throw new Error('Not signed in');

  try {
    const response = await instance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: account,
    });
    return response.accessToken;
  } catch (e) {
    console.warn('Silent token acquisition failed, redirecting:', e.message);
    await instance.acquireTokenRedirect({
      scopes: GRAPH_SCOPES,
      account: account,
    });
    return null;
  }
};

// ── Sign In (redirect flow) ─────────────────────────────────────────────────
window.signIn = async function () {
  try {
    const instance = getMsal();
    await instance.loginRedirect({ scopes: GRAPH_SCOPES });
  } catch (e) {
    console.error('Sign-in error:', e);
    alert('Sign-in failed: ' + e.message);
  }
};

// ── Sign Out ────────────────────────────────────────────────────────────────
window.signOut = function () {
  const instance = getMsal();
  instance.logoutRedirect({
    postLogoutRedirectUri: AUTH_CONFIG.redirectUri,
  });
};

// ── Show the dashboard after successful auth ────────────────────────────────
function showDashboard(account) {
  _currentAccount = account;
  document.getElementById('hdr-user').textContent = account.name || account.username || 'User';
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-loading').classList.add('active');

  const badge = document.getElementById('api-badge');
  if (badge) {
    badge.textContent = '🟢 Live SharePoint';
    badge.style.background = 'rgba(255,255,255,.18)';
  }

  // Notify app.js that auth is done
  if (typeof onSignedIn === 'function') onSignedIn(account);
}

// ── Handle MSAL redirect on page load (v2 API — no initialize() needed) ─────
(async function initAuth() {
  try {
    const instance = getMsal();

    // Handle the redirect response (if returning from login)
    const redirectResult = await instance.handleRedirectPromise();

    if (redirectResult && redirectResult.account) {
      instance.setActiveAccount(redirectResult.account);
      showDashboard(redirectResult.account);
      return;
    }

    // Check if already signed in (cached session)
    const accounts = instance.getAllAccounts();
    if (accounts.length > 0) {
      instance.setActiveAccount(accounts[0]);
      showDashboard(accounts[0]);
      return;
    }

    // Not signed in — stay on login page (default)
  } catch (e) {
    console.error('MSAL init error:', e);
    const note = document.querySelector('.login-note');
    if (note) note.textContent = 'Auth error: ' + e.message;
  }
})();
