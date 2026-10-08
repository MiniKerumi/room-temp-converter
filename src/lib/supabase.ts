import { createClient } from "@supabase/supabase-js";

const url = import.meta.env['VITE_SUPABASE_URL'] as string | undefined;
const key = import.meta.env['VITE_SUPABASE_ANON_KEY'] as string | undefined;

/** False until a backend is connected — the app must not call sign-in or data features then. */
export const supabaseConfigured = Boolean(url && key);

// Placeholder values keep the page from crashing when no backend is connected yet.
export const supabase = createClient(url || "https://not-configured.invalid", key || "not-configured");
