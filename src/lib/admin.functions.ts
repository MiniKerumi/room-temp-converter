import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";
import { requireSupabaseAuth } from "@/integrations/supabase/auth-middleware";

const accountInput = z.object({
  email: z.string().email(),
  full_name: z.string().min(1),
  password: z.string().min(8),
  role: z.enum(["staff", "admin"]),
});

async function getAdmin() {
  const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
  return supabaseAdmin;
}

/** Creates the very first admin account. Only works while no admin exists yet. */
export const bootstrapAdmin = createServerFn({ method: "POST" })
  .inputValidator((data) => accountInput.parse(data))
  .handler(async ({ data }) => {
    const admin = await getAdmin();
    const { data: existing, error: roleError } = await admin
      .from("user_roles")
      .select("id")
      .eq("role", "admin")
      .limit(1);
    if (roleError) throw new Error(roleError.message);
    if (existing && existing.length > 0) {
      throw new Error("An administrator already exists. Ask them to create your account.");
    }
    return createOrUpdateAccount(admin, data);
  });

/** Admin-only: create a staff/admin account or reset an existing one's password. */
export const manageAccount = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator((data) => accountInput.parse(data))
  .handler(async ({ data, context }) => {
    const { data: isAdmin, error: checkError } = await context.supabase.rpc("has_role", {
      _user_id: context.userId,
      _role: "admin",
    });
    if (checkError) throw new Error(checkError.message);
    if (!isAdmin) throw new Error("Only administrators can manage accounts.");
    const admin = await getAdmin();
    return createOrUpdateAccount(admin, data);
  });

type AdminClient = Awaited<ReturnType<typeof getAdmin>>;

async function createOrUpdateAccount(
  admin: AdminClient,
  data: z.infer<typeof accountInput>,
) {
  const email = data.email.trim().toLowerCase();

  // Find existing auth user by email (paginated list; fine for small teams).
  let userId: string | null = null;
  const { data: list, error: listError } = await admin.auth.admin.listUsers({ perPage: 1000 });
  if (listError) throw new Error(listError.message);
  const found = list.users.find((u) => u.email?.toLowerCase() === email);

  if (found) {
    const { error: updateError } = await admin.auth.admin.updateUserById(found.id, {
      password: data.password,
      email_confirm: true,
    });
    if (updateError) throw new Error(updateError.message);
    userId = found.id;
  } else {
    const { data: created, error: createError } = await admin.auth.admin.createUser({
      email,
      password: data.password,
      email_confirm: true,
    });
    if (createError) throw new Error(createError.message);
    userId = created.user.id;
  }

  const { error: profileError } = await admin.from("profiles").upsert({
    id: userId,
    email,
    full_name: data.full_name,
  });
  if (profileError) throw new Error(profileError.message);

  await admin.from("user_roles").delete().eq("user_id", userId);
  const { error: roleInsertError } = await admin
    .from("user_roles")
    .insert({ user_id: userId, role: data.role });
  if (roleInsertError) throw new Error(roleInsertError.message);

  return { ok: true };
}
