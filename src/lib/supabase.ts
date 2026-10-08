// Re-export the generated Lovable Cloud client so existing imports keep working.
export { supabase } from "@/integrations/supabase/client";

/** Lovable Cloud is connected, so sign-in and data features are available. */
export const supabaseConfigured = true;
