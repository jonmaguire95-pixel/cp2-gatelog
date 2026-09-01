import { createClient } from "@supabase/supabase-js";

// Public client credentials for the CP2 Gate Log database. These are safe to
// keep in code (they ship inside the built bundle either way) and they live
// here on purpose: the Vercel environment copy of the key picked up a stray
// character and silently broke every database call in production (Sep 1).
// The override variables below are for local testing only and are not set in
// production builds.
const SUPABASE_URL =
  import.meta.env.VITE_SUPABASE_URL_OVERRIDE ||
  "https://jaovecnuxuxmxngfyifb.supabase.co";
const SUPABASE_ANON_KEY =
  import.meta.env.VITE_SUPABASE_KEY_OVERRIDE ||
  "sb_publishable_ky4Sq8k4tgOZwVVwWGx6Mg_bCs5GyvT";

export const SHARED_PASSWORD = import.meta.env.VITE_SHARED_PASSWORD || "cp2gate2026";

export const supabase = SUPABASE_URL && SUPABASE_ANON_KEY
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : null;

export const supabaseReady = () => supabase !== null;
