import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const AdminToken = z.string().min(1, "جلسة الإدارة غير متاحة");

const CreateManagerSchema = z.object({
  email: z.string().email(),
  password: z.string().min(6),
  name: z.string().min(1),
  permissions: z.string(),
  accessToken: AdminToken,
});

const DeleteManagerSchema = z.object({
  managerId: z.string().min(1),
  accessToken: AdminToken,
});

function roleFromPermissions(permissions: string): "manager" | "viewer" {
  return permissions === "مدير كامل الصلاحيات" ? "manager" : "viewer";
}

async function getAdminClient() {
  const { supabaseAdmin } = await import("@/integrations/supabase/client.server");
  return supabaseAdmin;
}

async function assertAdminAccess(accessToken: string) {
  const supabaseAdmin = await getAdminClient();
  const { data: userData, error: userError } = await supabaseAdmin.auth.getUser(accessToken);
  if (userError || !userData.user) {
    throw new Error("انتهت جلسة الإدارة، يرجى تسجيل الدخول مرة أخرى");
  }

  const [{ data: roleRows, error: roleError }, { data: managerRow, error: managerError }] = await Promise.all([
    supabaseAdmin.from("user_roles").select("role").eq("user_id", userData.user.id),
    supabaseAdmin
      .from("admin_affairs_managers")
      .select("permissions")
      .eq("email", userData.user.email || "")
      .maybeSingle(),
  ]);

  if (roleError) throw new Error(roleError.message);
  if (managerError) throw new Error(managerError.message);

  const isOwnerOrManager = (roleRows || []).some(
    (row: any) => row.role === "owner" || row.role === "manager",
  );
  const hasFullManagerPermissions = managerRow?.permissions === "مدير كامل الصلاحيات";

  if (!isOwnerOrManager && !hasFullManagerPermissions) {
    throw new Error("ليس لديك صلاحية إدارة المستخدمين");
  }

  return userData.user;
}

async function createAuthUserWithRole(
  email: string,
  password: string,
  name: string,
  permissions: string,
  role: "manager" | "viewer",
) {
  const supabaseAdmin = await getAdminClient();
  let userId: string | null = null;

  try {
    const { data: created, error: createError } = await supabaseAdmin.auth.admin.createUser({
      email,
      password,
      email_confirm: true,
    });

    if (createError || !created?.user) {
      throw new Error(createError?.message || "تعذر إنشاء المستخدم");
    }

    userId = created.user.id;

    const { error: roleError } = await supabaseAdmin
      .from("user_roles")
      .insert([{ user_id: userId, role }]);
    if (roleError) throw new Error(roleError.message);

    // جدول المديرين في قاعدة بياناتك لا يحتوي على user_id، لذلك نربطه بالبريد الإلكتروني.
    const { error: managerError } = await supabaseAdmin
      .from("admin_affairs_managers")
      .insert([{ name, email, permissions }]);
    if (managerError) throw new Error(managerError.message);

    return { userId };
  } catch (error) {
    // تنظيف أي حساب جزئي إذا فشلت خطوة لاحقة.
    if (userId) {
      await supabaseAdmin.from("user_roles").delete().eq("user_id", userId);
      await supabaseAdmin.auth.admin.deleteUser(userId);
    }
    throw error;
  }
}

export const createManagerAccount = createServerFn({ method: "POST" })
  .validator((input: unknown) => CreateManagerSchema.parse(input))
  .handler(async ({ data }) => {
    await assertAdminAccess(data.accessToken);
    const role = roleFromPermissions(data.permissions);
    return createAuthUserWithRole(data.email, data.password, data.name, data.permissions, role);
  });

export const deleteManagerAccount = createServerFn({ method: "POST" })
  .validator((input: unknown) => DeleteManagerSchema.parse(input))
  .handler(async ({ data }) => {
    const requester = await assertAdminAccess(data.accessToken);
    const supabaseAdmin = await getAdminClient();

    const { data: manager, error: lookupError } = await supabaseAdmin
      .from("admin_affairs_managers")
      .select("email")
      .eq("id", data.managerId)
      .maybeSingle();

    if (lookupError) throw new Error(lookupError.message);
    if (!manager?.email) throw new Error("لم يتم العثور على المستخدم المطلوب");

    if (manager.email.toLowerCase() === (requester.email || "").toLowerCase()) {
      throw new Error("لا يمكن حذف المستخدم الذي سجل الدخول حاليًا");
    }

    // المستخدمون الحاليون قديمًا لا يملكون user_id في جدول المديرين، لذلك نبحث بالبريد.
    const { data: usersPage, error: listError } = await supabaseAdmin.auth.admin.listUsers({
      page: 1,
      perPage: 1000,
    });
    if (listError) throw new Error(listError.message);

    const authUser = (usersPage.users || []).find(
      (user: any) => (user.email || "").toLowerCase() === manager.email.toLowerCase(),
    );

    if (authUser) {
      const { error: authDeleteError } = await supabaseAdmin.auth.admin.deleteUser(authUser.id);
      if (authDeleteError) throw new Error(authDeleteError.message);
      await supabaseAdmin.from("user_roles").delete().eq("user_id", authUser.id);
    }

    const { error: managerDeleteError } = await supabaseAdmin
      .from("admin_affairs_managers")
      .delete()
      .eq("id", data.managerId);
    if (managerDeleteError) throw new Error(managerDeleteError.message);

    return { success: true, userId: authUser?.id || null };
  });
