import { describe, it, expect } from "vitest";
import worker, { findTokenName } from "../src/index.js";

const GOOD = "a".repeat(40);
const OTHER = "b".repeat(40);

function env(overrides = {}) {
  return {
    RD_SERVER: "rd.example.org",
    RD_KEY: "PUBKEY=",
    DEPLOY_TOKENS: `max:${GOOD}, helpdesk:${OTHER}`,
    LIMITER: { limit: async () => ({ success: true }) },
    ...overrides,
  };
}

function get(path = "/config", token, method = "GET") {
  const headers = { "cf-connecting-ip": "203.0.113.7" };
  if (token !== undefined) headers.authorization = `Bearer ${token}`;
  return new Request(`https://rd.tebin.pro${path}`, { method, headers });
}

describe("findTokenName", () => {
  it("returns the admin name for a listed token", async () => {
    expect(await findTokenName(OTHER, `max:${GOOD},helpdesk:${OTHER}`)).toBe("helpdesk");
  });
  it("rejects unknown, empty and short tokens", async () => {
    expect(await findTokenName("c".repeat(40), `max:${GOOD}`)).toBeNull();
    expect(await findTokenName("", `max:${GOOD}`)).toBeNull();
    expect(await findTokenName("short", "max:short")).toBeNull();
  });
  it("rejects everything when the secret is missing", async () => {
    expect(await findTokenName(GOOD, undefined)).toBeNull();
  });
});

describe("fetch", () => {
  it("returns server and key for a valid token, never cached", async () => {
    const res = await worker.fetch(get("/config", GOOD), env());
    expect(res.status).toBe(200);
    expect(res.headers.get("cache-control")).toBe("no-store");
    expect(await res.json()).toEqual({ server: "rd.example.org", key: "PUBKEY=" });
  });
  it("returns 401 without the secrets for a wrong or missing token", async () => {
    for (const token of ["x".repeat(40), undefined]) {
      const res = await worker.fetch(get("/config", token), env());
      expect(res.status).toBe(401);
      expect(await res.text()).not.toContain("PUBKEY");
    }
  });
  it("returns 429 when the limiter refuses, before checking the token", async () => {
    const res = await worker.fetch(get("/config", GOOD), env({ LIMITER: { limit: async () => ({ success: false }) } }));
    expect(res.status).toBe(429);
  });
  it("fails open when the limiter throws", async () => {
    const res = await worker.fetch(get("/config", GOOD), env({ LIMITER: { limit: async () => { throw new Error("down"); } } }));
    expect(res.status).toBe(200);
  });
  it("returns 500 when server or key secret is empty", async () => {
    const res = await worker.fetch(get("/config", GOOD), env({ RD_KEY: "" }));
    expect(res.status).toBe(500);
  });
  it("answers 404 / 405 for other paths and methods", async () => {
    expect((await worker.fetch(get("/", GOOD), env())).status).toBe(404);
    expect((await worker.fetch(get("/config", GOOD, "POST"), env())).status).toBe(405);
  });
});
