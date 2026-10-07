/*
# Tighten inspection app permissions

1. Security Changes
- Remove all direct table privileges from the unauthenticated `anon` role.
- Remove public and unauthenticated execution of helper functions.
- Keep signed-in access limited to the row-level policies created for staff and administrators.

2. Important Notes
- Supabase Auth triggers continue to use the protected `handle_new_user` function internally.
- Account management continues through the deployed administrator-only Edge Function.
- No application data is deleted or changed.
*/

REVOKE ALL ON TABLE public.profiles FROM anon;
REVOKE ALL ON TABLE public.inspection_submissions FROM anon;
REVOKE EXECUTE ON FUNCTION public.is_admin() FROM PUBLIC;
REVOKE EXECUTE ON FUNCTION public.is_admin() FROM anon;
GRANT EXECUTE ON FUNCTION public.is_admin() TO authenticated;
REVOKE EXECUTE ON FUNCTION public.handle_new_user() FROM PUBLIC;
REVOKE EXECUTE ON FUNCTION public.handle_new_user() FROM anon;
REVOKE EXECUTE ON FUNCTION public.handle_new_user() FROM authenticated;
