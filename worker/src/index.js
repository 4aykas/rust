// rd-config: hands the RustDesk server address and key to rust.ps1, only for a
// valid deploy token. Values live in Worker secrets, never in the public repo.
//
// Secrets:
//   RD_SERVER      e.g. rd.example.org
//   RD_KEY         public key of hbbs (id_ed25519.pub)
//   DEPLOY_TOKENS  comma-separated "name:token" pairs, one per admin, e.g.
//                  "max:Xy...,helpdesk:Ab..." (token: 32+ random chars)
//
// The token is the security boundary; the rate limiter only caps noise.

const encoder = new TextEncoder();

async function sha256(text) {
  return new Uint8Array(await crypto.subtle.digest("SHA-256", encoder.encode(text)));
}

// Compares fixed-length digests without an early exit.
function sameBytes(a, b) {
  let diff = a.length ^ b.length;
  for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i];
  return diff === 0;
}

// Returns the admin name for a token, or null. Checks every entry so the time
// taken does not tell which entry nearly matched.
export async function findTokenName(token, tokensSecret) {
  if (!token) return null;
  const given = await sha256(token);
  let found = null;
  for (const entry of (tokensSecret ?? "").split(",")) {
    const at = entry.indexOf(":");
    if (at < 1) continue;
    const name = entry.slice(0, at).trim();
    const value = entry.slice(at + 1).trim();
    if (value.length < 32) continue;
    if (sameBytes(given, await sha256(value)) && found === null) found = name;
  }
  return found;
}

function json(body, status) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "content-type": "application/json",
      "cache-control": "no-store",
      "x-content-type-options": "nosniff",
    },
  });
}

function log(event, fields) {
  console.log(JSON.stringify({ event, ...fields }));
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname !== "/config") return json({ error: "not found" }, 404);
    if (request.method !== "GET") return json({ error: "method not allowed" }, 405);

    const ip = request.headers.get("cf-connecting-ip") ?? "unknown";

    // Fail open: a broken limiter must not block deployments.
    try {
      const { success } = await env.LIMITER.limit({ key: ip });
      log("ratelimit", { ip, success });
      if (!success) return json({ error: "too many attempts" }, 429);
    } catch (err) {
      log("ratelimit_error", { ip, error: String(err) });
    }

    const auth = request.headers.get("authorization") ?? "";
    const token = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";
    const name = await findTokenName(token, env.DEPLOY_TOKENS);
    if (!name) {
      log("denied", { ip });
      return json({ error: "invalid token" }, 401);
    }
    if (!env.RD_SERVER || !env.RD_KEY) {
      log("misconfigured", { ip, name });
      return json({ error: "service not configured" }, 500);
    }
    log("granted", { ip, name });
    return json({ server: env.RD_SERVER, key: env.RD_KEY }, 200);
  },
};
