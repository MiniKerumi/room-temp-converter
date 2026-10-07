import { createClient } from "npm:@supabase/supabase-js@2.57.4";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Client-Info, Apikey",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { status: 200, headers: corsHeaders });
  try {
    const authHeader = req.headers.get("Authorization");
    if (!authHeader) return new Response(JSON.stringify({ error: "Not authorized" }), { status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" } });
    const url = Deno.env.get("SUPABASE_URL")!;
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
    const caller = createClient(url, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: authHeader } } });
    const admin = createClient(url, serviceKey);
    const { data: userData } = await caller.auth.getUser();
    if (!userData.user) return new Response(JSON.stringify({ error: "Not authorized" }), { status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" } });
    const { data: profile } = await admin.from("profiles").select("role").eq("id", userData.user.id).maybeSingle();
    if (profile?.role !== "admin") return new Response(JSON.stringify({ error: "Not authorized" }), { status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" } });
    const body = await req.json() as { email?: string; full_name?: string; password?: string; role?: "admin" | "staff" };
    const email = body.email?.trim().toLowerCase();
    const password = body.password ?? "";
    if (!email || !email.includes("@") || password.length < 8 || !body.full_name?.trim() || !body.role) return new Response(JSON.stringify({ error: "Complete all account fields and use a password with at least 8 characters." }), { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } });
    const listed = await admin.auth.admin.listUsers({ page: 1, perPage: 1000 });
    if (listed.error) throw listed.error;
    const existing = listed.data.users.find((user) => user.email?.toLowerCase() === email);
    if (existing) {
      const updated = await admin.auth.admin.updateUserById(existing.id, { password, user_metadata: { full_name: body.full_name } });
      if (updated.error) throw updated.error;
      const { error } = await admin.from("profiles").update({ email, full_name: body.full_name, role: body.role, updated_at: new Date().toISOString() }).eq("id", existing.id);
      if (error) throw error;
      return new Response(JSON.stringify({ message: "Account updated." }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }
    const created = await admin.auth.admin.createUser({ email, password, email_confirm: true, user_metadata: { full_name: body.full_name } });
    if (created.error || !created.data.user) throw created.error ?? new Error("Account could not be created");
    const { error } = await admin.from("profiles").update({ full_name: body.full_name, role: body.role }).eq("id", created.data.user.id);
    if (error) throw error;
    return new Response(JSON.stringify({ message: "Account created." }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
  } catch (error) {
    console.error(error);
    return new Response(JSON.stringify({ error: "Account could not be saved." }), { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } });
  }
});
