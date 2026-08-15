import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const CreateManagerSchema = z.object({
  email: z.string().email(),
  password: z.string().min(6),
  name: z.string().min(1),
  permissions: z.string(),
});

function roleFromPermissions(permissions: string): "manager" | "viewer" {
  return permissions === "مدير كامل الصلاحيات" ? "manager" : "viewer";
}

async function createAuthUserWithRole(
  email: string,
  password: string,
  name: string,
  permissions: string,
  role: "manager" | "viewer",
) {
  const { supabaseAdmin } = await import("@/integrations/supabase/client.server");

  // ينشئ المستخدم فعلياً في نظام Supabase Auth (وليس في جدول عادي).
  const { data: created, error: createError } = await supabaseAdmin.auth.admin.createUser({
    email,
    password,
    email_confirm: true,
  });
  if (createError || !created?.user) {
    throw new Error(createError?.message || "تعذر إنشاء المستخدم");
  }

  const userId = created.user.id;

  const { error: roleError } = await supabaseAdmin.from("user_roles").insert([{ user_id: userId, role }]);
  if (roleError) throw new Error(roleError.message);

  const { error: mgrError } = await supabaseAdmin
    .from("admin_affairs_managers")
    .insert([{ user_id: userId, name, email, permissions }]);
  if (mgrError) throw new Error(mgrError.message);

  return { userId };
}

// إنشاء الحسابات يتم فقط من شاشة المستخدمين بعد مصادقة Owner/مدير.
// لا ننشئ حسابًا ذا صلاحيات عالية داخل مسار تسجيل الدخول نفسه.

// يُستدعى من شاشة "إضافة مستخدم" في تبويب المستخدمين.
export const createManagerAccount = createServerFn({ method: "POST" })
  .validator((input: unknown) => CreateManagerSchema.parse(input))
  .handler(async ({ data }) => {
    const role = roleFromPermissions(data.permissions);
    return createAuthUserWithRole(data.email, data.password, data.name, data.permissions, role);
  });
