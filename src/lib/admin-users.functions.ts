// ضع هذا الملف في نفس المسار: src/lib/admin-users.functions.ts
import { supabase } from "@/integrations/supabase/client";

function genId(): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function createUserWithRole({
  email,
  password,
  name,
  permissions,
  role,
}: {
  email: string;
  password: string;
  name: string;
  permissions: string;
  role: "owner" | "manager" | "viewer";
}) {
  const { data: existing } = await supabase.from("auth_users").select("*").eq("email", email).maybeSingle();
  if (existing) {
    throw new Error("هذا البريد الإلكتروني مسجل بالفعل");
  }

  const userId = genId();
  await supabase.from("auth_users").insert([{ id: userId, email, password }]);
  await supabase.from("user_roles").insert([{ id: genId(), user_id: userId, role }]);
  await supabase.from("admin_affairs_managers").insert([
    { id: genId(), user_id: userId, name, email, password, permissions },
  ]);
  return { userId };
}

// يُستدعى عند كل محاولة تسجيل دخول. إذا لم يوجد مالك بعد، يُنشئ الحساب
// المُدخل كأول مالك للنظام. إذا كان هناك مالك بالفعل، لا يفعل شيئاً.
export async function bootstrapOwner({
  data,
}: {
  data: { email: string; password: string };
}) {
  const { data: owners } = await supabase.from("user_roles").select("*").eq("role", "owner");
  if (owners && owners.length > 0) {
    return { alreadyExists: true };
  }
  return createUserWithRole({
    email: data.email,
    password: data.password,
    name: data.email.split("@")[0],
    permissions: "مدير كامل الصلاحيات",
    role: "owner",
  });
}

// يُستدعى من شاشة "إضافة مستخدم" في تبويب المستخدمين.
export async function createManagerAccount({
  data,
}: {
  data: { email: string; password: string; name: string; permissions: string };
}) {
  const role = data.permissions === "مدير كامل الصلاحيات" ? "manager" : "viewer";
  return createUserWithRole({
    email: data.email,
    password: data.password,
    name: data.name,
    permissions: data.permissions,
    role,
  });
}
