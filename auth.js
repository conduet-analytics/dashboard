/**
 * auth.js
 * -------
 * MSAL v2 authentication for SharePoint-hosted SPA.
 * Authenticates via Azure AD — the user's token is used to call
 * Microsoft Graph directly from the browser to read SharePoint Excel data.
 */

'use strict';

// ── Guard: check MSAL loaded ─────────────────────────────────────────────────
if (typeof msal === 'undefined') {
  console.error('MSAL library failed to load from CDN.');
  document.addEventListener('DOMContentLoaded', () => {
    const note = document.querySelector('.login-note');
    if (note) {
      note.style.color = '#c0392b';
      note.innerHTML = 'MSAL library failed to load. Check browser console (F12 → Network tab) for blocked scripts.';
    }
  });
  window.signIn = function () {
    alert('MSAL library did not load. Open DevTools (F12) → Network tab → reload the page and look for a failed request to alcdn.msauth.net');
  };
  window.signOut = function () {};
  window.getToken = function () { return null; };
} else {

// ── Azure AD app registration ────────────────────────────────────────────────
const AUTH_CONFIG = {
  clientId:  'fb723393-bd9d-4052-9f18-e9b65a6174c4',
  tenantId:  '3acfc3b8-ea64-4ce1-8910-f91e9c67a3fc',
  get redirectUri() {
    // Works automatically on any host (SharePoint, localhost, etc.)
    return window.location.origin + window.location.pathname;
  },
};

// Scopes matching what IT approved
const GRAPH_SCOPES = [
  'User.Read',
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All',
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

  if (typeof onSignedIn === 'function') onSignedIn(account);
}

// ── Handle MSAL redirect on page load ────────────────────────────────────────
(async function initAuth() {
  try {
    const instance = getMsal();
    const redirectResult = await instance.handleRedirectPromise();

    if (redirectResult && redirectResult.account) {
      instance.setActiveAccount(redirectResult.account);
      showDashboard(redirectResult.account);
      return;
    }

    const accounts = instance.getAllAccounts();
    if (accounts.length > 0) {
      instance.setActiveAccount(accounts[0]);
      showDashboard(accounts[0]);
      return;
    }
  } catch (e) {
    console.error('MSAL init error:', e);
    const note = document.querySelector('.login-note');
    if (note) note.textContent = 'Auth error: ' + e.message;
  }
})();

} // end of else (MSAL loaded successfully)
