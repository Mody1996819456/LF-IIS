// ==============================================================
// استبدال محلي لعميل Supabase — كل البيانات تُخزَّن في localStorage
// ==============================================================
// ملاحظات مهمة:
// - البيانات مرتبطة بهذا المتصفح/الجهاز فقط، ولا تُشارك بين الأجهزة.
// - لا يوجد تشفير حقيقي لكلمات المرور (تُخزَّن كنص عادي).
// - مساحة التخزين محدودة (~5-10 ميجابايت حسب المتصفح).
// ضع هذا الملف في نفس المسار: src/integrations/supabase/client.ts

type FilterOp = "eq" | "neq";
type Filter = { col: string; val: any; op: FilterOp };

function loadTable(table: string): any[] {
  try {
    const raw = localStorage.getItem(`localdb_${table}`);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function saveTable(table: string, rows: any[]) {
  try {
    localStorage.setItem(`localdb_${table}`, JSON.stringify(rows));
  } catch (e) {
    console.error(`تعذر حفظ الجدول ${table} في localStorage`, e);
  }
}

function genId(): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function compareValues(a: any, b: any): number {
  if (a == null && b == null) return 0;
  if (a == null) return -1;
  if (b == null) return 1;
  const aDate = typeof a === "string" && !isNaN(Date.parse(a)) ? Date.parse(a) : null;
  const bDate = typeof b === "string" && !isNaN(Date.parse(b)) ? Date.parse(b) : null;
  if (aDate !== null && bDate !== null) return aDate - bDate;
  if (a < b) return -1;
  if (a > b) return 1;
  return 0;
}

class LocalQueryBuilder {
  private table: string;
  private operation: "select" | "insert" | "update" | "delete" | null = null;
  private payload: any = null;
  private filters: Filter[] = [];
  private orderCol: string | null = null;
  private orderAsc = true;
  private wantSingle = false;
  private wantReturnAfterWrite = false;

  constructor(table: string) {
    this.table = table;
  }

  select(_cols?: string) {
    if (!this.operation) this.operation = "select";
    else this.wantReturnAfterWrite = true;
    return this;
  }

  insert(rows: any[]) {
    this.operation = "insert";
    this.payload = rows;
    return this;
  }

  update(obj: any) {
    this.operation = "update";
    this.payload = obj;
    return this;
  }

  delete() {
    this.operation = "delete";
    return this;
  }

  eq(col: string, val: any) {
    this.filters.push({ col, val, op: "eq" });
    return this;
  }

  neq(col: string, val: any) {
    this.filters.push({ col, val, op: "neq" });
    return this;
  }

  order(col: string, opts?: { ascending?: boolean }) {
    this.orderCol = col;
    this.orderAsc = opts?.ascending !== false;
    return this;
  }

  maybeSingle() {
    this.wantSingle = true;
    return this;
  }

  private matchesFilters(row: any): boolean {
    return this.filters.every(f => {
      if (f.op === "eq") return row[f.col] === f.val;
      if (f.op === "neq") return row[f.col] !== f.val;
      return true;
    });
  }

  private execute(): { data: any; error: any } {
    try {
      const rows = loadTable(this.table);

      if (this.operation === "select" || this.operation === null) {
        let result = rows.filter(r => this.matchesFilters(r));
        if (this.orderCol) {
          const col = this.orderCol;
          result = [...result].sort((a, b) => {
            const c = compareValues(a[col], b[col]);
            return this.orderAsc ? c : -c;
          });
        }
        if (this.wantSingle) {
          return { data: result[0] ?? null, error: null };
        }
        return { data: result, error: null };
      }

      if (this.operation === "insert") {
        const toInsert = (this.payload as any[]).map(r => ({ ...r, id: r.id || genId() }));
        saveTable(this.table, [...rows, ...toInsert]);
        return { data: this.wantReturnAfterWrite ? toInsert : null, error: null };
      }

      if (this.operation === "update") {
        const matched: any[] = [];
        const updated = rows.map(r => {
          if (this.matchesFilters(r)) {
            const merged = { ...r, ...this.payload };
            matched.push(merged);
            return merged;
          }
          return r;
        });
        saveTable(this.table, updated);
        return { data: this.wantReturnAfterWrite ? matched : null, error: null };
      }

      if (this.operation === "delete") {
        const toKeep = rows.filter(r => !this.matchesFilters(r));
        saveTable(this.table, toKeep);
        return { data: null, error: null };
      }

      return { data: null, error: null };
    } catch (e: any) {
      return { data: null, error: { message: e?.message || "خطأ في التخزين المحلي" } };
    }
  }

  // يسمح باستخدام await مباشرة على السلسلة، تماماً كما في supabase-js
  then(onFulfilled: any, onRejected?: any) {
    const result = this.execute();
    return Promise.resolve(result).then(onFulfilled, onRejected);
  }
}

// ---- محاكاة نظام تسجيل الدخول (محلي بالكامل) ----
type AuthListener = (event: string, session: any) => void;
const authListeners: AuthListener[] = [];

function getSessionFromStorage() {
  try {
    const raw = localStorage.getItem("local_auth_session");
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function setSessionInStorage(session: any) {
  if (session) localStorage.setItem("local_auth_session", JSON.stringify(session));
  else localStorage.removeItem("local_auth_session");
}

function notifyAuthChange(event: string, session: any) {
  authListeners.forEach(l => l(event, session));
}

export const supabase = {
  from(table: string) {
    return new LocalQueryBuilder(table);
  },
  auth: {
    async signInWithPassword({ email, password }: { email: string; password: string }) {
      const users = loadTable("auth_users");
      const user = users.find((u: any) => u.email === email && u.password === password);
      if (!user) {
        return { data: { session: null }, error: { message: "بيانات الدخول غير صحيحة" } };
      }
      const session = { user: { id: user.id, email: user.email } };
      setSessionInStorage(session);
      notifyAuthChange("SIGNED_IN", session);
      return { data: { session }, error: null };
    },
    async signOut() {
      setSessionInStorage(null);
      notifyAuthChange("SIGNED_OUT", null);
      return { error: null };
    },
    async getSession() {
      return { data: { session: getSessionFromStorage() } };
    },
    onAuthStateChange(callback: AuthListener) {
      authListeners.push(callback);
      return {
        data: {
          subscription: {
            unsubscribe() {
              const idx = authListeners.indexOf(callback);
              if (idx >= 0) authListeners.splice(idx, 1);
            },
          },
        },
      };
    },
  },
};

// ==============================================================
// تهيئة تلقائية: إنشاء حساب المالك المطلوب عند أول تحميل للتطبيق
// (لن يُنشأ إلا مرة واحدة، وإذا كان هناك مالك بالفعل لن يُكرر)
// ==============================================================
(function seedOwnerAccount() {
  const OWNER_EMAIL = "heshamarab2@gmail.com";
  const OWNER_PASSWORD = "2026";
  try {
    const roles = loadTable("user_roles");
    const hasOwner = roles.some((r: any) => r.role === "owner");
    if (hasOwner) return;

    const users = loadTable("auth_users");
    if (users.some((u: any) => u.email === OWNER_EMAIL)) return;

    const userId = genId();
    saveTable("auth_users", [...users, { id: userId, email: OWNER_EMAIL, password: OWNER_PASSWORD }]);
    saveTable("user_roles", [...roles, { id: genId(), user_id: userId, role: "owner" }]);

    const managers = loadTable("admin_affairs_managers");
    saveTable("admin_affairs_managers", [
      ...managers,
      {
        id: genId(),
        user_id: userId,
        name: "Hesham",
        email: OWNER_EMAIL,
        password: OWNER_PASSWORD,
        permissions: "مدير كامل الصلاحيات",
      },
    ]);
  } catch {
    // localStorage غير متاح (مثلاً أثناء SSR) — تجاهل بأمان.
  }
})();
