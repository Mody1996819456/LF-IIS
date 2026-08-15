import { createClient } from "@supabase/supabase-js";
import type { Database } from "./types";

function isNewSupabaseApiKey(value: string): boolean {
  return value.startsWith("sb_publishable_") || value.startsWith("sb_secret_");
}

function createSupabaseFetch(supabaseKey: string): typeof fetch {
  return (input, init) => {
    const headers = new Headers(
      typeof Request !== "undefined" && input instanceof Request ? input.headers : undefined,
    );
    if (init?.headers) {
      new Headers(init.headers).forEach((value, key) => headers.set(key, value));
    }
    if (isNewSupabaseApiKey(supabaseKey) && headers.get("Authorization") === `Bearer ${supabaseKey}`) {
      headers.delete("Authorization");
    }
    headers.set("apikey", supabaseKey);
    return fetch(input, { ...init, headers });
  };
}

const SUPABASE_URL = String(import.meta.env.VITE_SUPABASE_URL || "").trim();
const SUPABASE_PUBLISHABLE_KEY = String(import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || "").trim();

export const isSupabaseConfigured = Boolean(SUPABASE_URL && SUPABASE_PUBLISHABLE_KEY);

if (!isSupabaseConfigured) {
  console.warn(
    "[Supabase] Missing VITE_SUPABASE_URL / VITE_SUPABASE_PUBLISHABLE_KEY. " +
      "The app will show a configuration screen instead of crashing during SSR.",
  );
}

// Supabase requires a syntactically valid URL even when the app is being previewed
// without its deployment environment. All network calls are short-circuited in that
// mode, while production still uses the real configured client and fetch wrapper.
const clientUrl = SUPABASE_URL || "https://supabase-not-configured.invalid";
const clientKey = SUPABASE_PUBLISHABLE_KEY || "supabase-not-configured";
const fallbackFetch: typeof fetch = async () =>
  new Response(JSON.stringify({ message: "Supabase is not configured" }), {
    status: 503,
    headers: { "Content-Type": "application/json" },
  });

export const supabase = createClient<Database>(clientUrl, clientKey, {
  global: {
    fetch: isSupabaseConfigured ? createSupabaseFetch(clientKey) : fallbackFetch,
  },
});
