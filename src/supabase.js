import { createClient } from "@supabase/supabase-js";

// Database connection settings. The URL and publishable key are public by
// design (they ship inside the page either way). Access control lives in the
// database itself: row level security only answers to a signed in account with
// a confirmed @mmrgrp.com email. The override variables are for local testing.
const SUPABASE_URL =
  import.meta.env.VITE_SUPABASE_URL_OVERRIDE ||
  "https://jaovecnuxuxmxngfyifb.supabase.co";
const SUPABASE_ANON_KEY =
  import.meta.env.VITE_SUPABASE_KEY_OVERRIDE ||
  "sb_publishable_ky4Sq8k4tgOZwVVwWGx6Mg_bCs5GyvT";

// Accounts are limited to MMR work email. The server enforces the same rule,
// so this is only here to give a clear message before anything is sent.
export const WORK_DOMAIN = "mmrgrp.com";

// Lowercases and trims. If someone types just the name part (jane.doe), the
// MMR domain is added for them.
export function normalizeEmail(value) {
  const v = String(value || "").trim().toLowerCase();
  return v && !v.includes("@") ? `${v}@${WORK_DOMAIN}` : v;
}

export function isWorkEmail(value) {
  return /^[^@\s]+@mmrgrp\.com$/.test(normalizeEmail(value));
}

// Email links (confirm a new account, reset a password) come back to the site
// with the sign in details after the # in the address. Read them once and wipe
// them from the address bar and from this history entry right away, before the
// page does anything else.
function takeLinkFromAddress() {
  if (typeof window === "undefined") return null;
  const { hash, search, pathname } = window.location;
  const fromHash = new URLSearchParams(hash.charAt(0) === "#" ? hash.slice(1) : hash);
  const fromQuery = new URLSearchParams(search);
  const pick = k => fromHash.get(k) || fromQuery.get(k) || "";
  const accessToken = fromHash.get("access_token") || "";
  const refreshToken = fromHash.get("refresh_token") || "";
  const errorCode = pick("error_code") || pick("error");
  const errorText = pick("error_description");
  if (!accessToken && !errorCode && !errorText) return null;

  ["error", "error_code", "error_description"].forEach(k => fromQuery.delete(k));
  const rest = fromQuery.toString();
  try {
    window.history.replaceState(window.history.state, "", pathname + (rest ? `?${rest}` : ""));
  } catch (_) {
    window.location.hash = "";
  }

  const expired = /expired|access_denied|invite_not_found/i.test(errorCode) || /expired|invalid/i.test(errorText);
  return {
    accessToken,
    refreshToken,
    type: fromHash.get("type") || "",
    error: accessToken ? null : expired ? "link-expired" : "link-failed",
  };
}

const arrivingLink = takeLinkFromAddress();

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    // The session lives in this browser tab only; closing the tab signs out.
    storage: typeof window !== "undefined" ? window.sessionStorage : undefined,
    persistSession: true,
    autoRefreshToken: true,
    // Email links are handled above so their tokens never linger in the
    // address bar. The implicit flow lets a link requested on a computer be
    // opened on a phone (or the other way around) and still work.
    detectSessionInUrl: false,
    flowType: "implicit",
  },
});

export const supabaseReady = () => supabase !== null;

// Where email links send people back to: this site's own address.
export function siteAddress() {
  return `${window.location.origin}/`;
}

// After arriving from an email link the person still has to choose a password.
// This remembers that for the current tab, so a page refresh doesn't skip it.
const STEP_KEY = "cp2gl-password-step";

export function readPasswordStep() {
  try {
    const v = window.sessionStorage.getItem(STEP_KEY);
    return v === "create" || v === "reset" ? v : null;
  } catch (_) {
    return null;
  }
}

export function writePasswordStep(step) {
  try {
    if (step) window.sessionStorage.setItem(STEP_KEY, step);
    else window.sessionStorage.removeItem(STEP_KEY);
  } catch (_) { /* storage blocked; the step just won't survive a refresh */ }
}

// Runs once when the page loads. Finishes an email link sign in if the page was
// opened from one, otherwise picks up an existing session for this tab.
// Resolves to { session, step, notice } and never rejects.
export const authStart = (async () => {
  const link = arrivingLink;
  try {
    if (link && link.accessToken && link.refreshToken) {
      const { data, error } = await supabase.auth.setSession({
        access_token: link.accessToken,
        refresh_token: link.refreshToken,
      });
      if (error || !data || !data.session) {
        writePasswordStep(null);
        return { session: null, step: null, notice: "link-failed" };
      }
      const meta = (data.session.user && data.session.user.user_metadata) || {};
      const step =
        link.type === "recovery" ? "reset" :
        link.type === "signup" || link.type === "invite" ? "create" :
        meta.password_set ? null : "create";
      writePasswordStep(step);
      return { session: data.session, step, notice: null };
    }
    const { data } = await supabase.auth.getSession();
    const session = (data && data.session) || null;
    return {
      session,
      step: session ? readPasswordStep() : null,
      notice: link && link.error ? link.error : null,
    };
  } catch (_) {
    return { session: null, step: null, notice: link ? "link-failed" : null };
  }
})();
