/*
# Create secure inspection app foundation

1. New Tables
- `profiles`: one row per signed-in account, with display name and immutable role (`admin` or `staff`).
- `inspection_submissions`: saved delivery vehicle inspection forms, owned by the signed-in creator and stored as structured JSON for accurate reproduction and export.

2. Authentication and Roles
- Uses Supabase Auth's built-in `auth.users` accounts; no custom password table is created.
- New accounts automatically receive a `staff` profile unless an administrator creates them through the protected admin function.
- The role is stored in `profiles.role` and is never accepted from the public browser during account creation.

3. Security
- Row-level security is enabled on every new table.
- Staff can access only their own inspection records and profile.
- Administrators can review and manage all inspection records.
- Account roles and ownership fields are protected from browser tampering.
- A server-enforced helper function is used for administrator checks.

4. Important Notes
- Passwords remain inside Supabase Auth and are never stored in application tables.
- The inspection payload is JSON because the form contains many related checkbox, remark, signature, and vehicle fields that must be preserved together for document export.
- Policies are split into separate SELECT, INSERT, UPDATE, and DELETE rules.
*/

CREATE TABLE IF NOT EXISTS public.profiles (
  id uuid PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  email text NOT NULL,
  full_name text NOT NULL DEFAULT '',
  role text NOT NULL DEFAULT 'staff' CHECK (role IN ('admin', 'staff')),
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.inspection_submissions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  created_by uuid NOT NULL DEFAULT auth.uid() REFERENCES auth.users(id) ON DELETE CASCADE,
  form_data jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS inspection_submissions_created_by_idx
  ON public.inspection_submissions (created_by);
CREATE INDEX IF NOT EXISTS inspection_submissions_created_at_idx
  ON public.inspection_submissions (created_at DESC);

ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.inspection_submissions ENABLE ROW LEVEL SECURITY;

CREATE OR REPLACE FUNCTION public.is_admin()
RETURNS boolean
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = public
AS $$
  SELECT EXISTS (
    SELECT 1 FROM public.profiles
    WHERE id = auth.uid() AND role = 'admin'
  );
$$;

REVOKE EXECUTE ON FUNCTION public.is_admin() FROM anon;
GRANT EXECUTE ON FUNCTION public.is_admin() TO authenticated;

CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS trigger
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
  INSERT INTO public.profiles (id, email, full_name)
  VALUES (NEW.id, NEW.email, COALESCE(NEW.raw_user_meta_data ->> 'full_name', ''))
  ON CONFLICT (id) DO UPDATE SET email = EXCLUDED.email;
  RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

DROP POLICY IF EXISTS "profiles_select_self_or_admin" ON public.profiles;
CREATE POLICY "profiles_select_self_or_admin"
  ON public.profiles FOR SELECT TO authenticated
  USING (id = auth.uid() OR public.is_admin());

DROP POLICY IF EXISTS "profiles_insert_self" ON public.profiles;
CREATE POLICY "profiles_insert_self"
  ON public.profiles FOR INSERT TO authenticated
  WITH CHECK (id = auth.uid() AND role = 'staff');

DROP POLICY IF EXISTS "profiles_update_self" ON public.profiles;
CREATE POLICY "profiles_update_self"
  ON public.profiles FOR UPDATE TO authenticated
  USING (id = auth.uid())
  WITH CHECK (id = auth.uid() AND role = 'staff');

DROP POLICY IF EXISTS "profiles_delete_self" ON public.profiles;
CREATE POLICY "profiles_delete_self"
  ON public.profiles FOR DELETE TO authenticated
  USING (false);

DROP POLICY IF EXISTS "inspections_select_own_or_admin" ON public.inspection_submissions;
CREATE POLICY "inspections_select_own_or_admin"
  ON public.inspection_submissions FOR SELECT TO authenticated
  USING (created_by = auth.uid() OR public.is_admin());

DROP POLICY IF EXISTS "inspections_insert_own" ON public.inspection_submissions;
CREATE POLICY "inspections_insert_own"
  ON public.inspection_submissions FOR INSERT TO authenticated
  WITH CHECK (created_by = auth.uid());

DROP POLICY IF EXISTS "inspections_update_own_or_admin" ON public.inspection_submissions;
CREATE POLICY "inspections_update_own_or_admin"
  ON public.inspection_submissions FOR UPDATE TO authenticated
  USING (created_by = auth.uid() OR public.is_admin())
  WITH CHECK (created_by = auth.uid() OR public.is_admin());

DROP POLICY IF EXISTS "inspections_delete_own_or_admin" ON public.inspection_submissions;
CREATE POLICY "inspections_delete_own_or_admin"
  ON public.inspection_submissions FOR DELETE TO authenticated
  USING (created_by = auth.uid() OR public.is_admin());

REVOKE INSERT, UPDATE, DELETE ON public.profiles FROM authenticated;
GRANT SELECT ON public.profiles TO authenticated;
REVOKE INSERT, UPDATE, DELETE ON public.inspection_submissions FROM authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON public.inspection_submissions TO authenticated;
