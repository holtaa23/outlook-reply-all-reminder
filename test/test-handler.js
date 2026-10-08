// Harness: runs commands.js against a stubbed Office / MSAL / Graph.
//
// Covers both data sources. Graph is the real one; the body parser is only a
// fallback, so several tests assert that Graph wins and that every way Graph
// can fail still degrades safely.
const fs = require("fs");
const vm = require("vm");
const path = require("path");

const SRC = process.argv[2] || path.join(__dirname, "..", "commands.js");
const code = fs.readFileSync(SRC, "utf8");

const ME = { email: "aaron@bex.com", name: "Aaron H" };
const CLIENT_ID = "11111111-2222-3333-4444-555555555555";

const R = (name, email) => ({ displayName: name, emailAddress: email });
const G = (name, address) => ({ emailAddress: { name, address } });

function makeSandbox(s) {
  const ok = (value) => (cb) => cb({ status: "succeeded", value });

  const Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Text: "text" },
    actions: { associate: () => {} },
    context: {
      mailbox: {
        userProfile: { emailAddress: ME.email, displayName: ME.name },
        item: {
          conversationId: s.conversationId === null ? null : "conv-1",
          getComposeTypeAsync: ok({ composeType: s.composeType }),
          body: { getAsync: (t, cb) => cb({ status: "succeeded", value: s.body || "" }) },
          to: { getAsync: ok(s.to || []) },
          cc: { getAsync: ok(s.cc || []) },
          bcc: { getAsync: ok(s.bcc || []) },
        },
      },
    },
  };

  // s.msal === false simulates classic Outlook's JavaScript-only runtime,
  // where the library never loads. s.tokenFails simulates no cached session.
  const msal =
    s.msal === false
      ? undefined
      : {
          createNestablePublicClientApplication: async () => ({
            getAllAccounts: () => [{ username: ME.email }],
            acquireTokenSilent: async () => {
              if (s.tokenFails) throw new Error("interaction_required");
              return { accessToken: "fake-token" };
            },
          }),
        };

  const fetch = async () => {
    if (s.graphFails) throw new Error("network");
    if (s.graphStatus && s.graphStatus !== 200) return { ok: false, status: s.graphStatus };
    if (s.graphHangs) return new Promise(() => {}); // never settles
    return { ok: true, status: 200, json: async () => ({ value: s.graphMessages || [] }) };
  };

  return { Office, msal, fetch, console, setTimeout, clearTimeout };
}

function run(scenario) {
  const sandbox = makeSandbox(scenario);
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox);

  sandbox.CONFIG.clientId =
    scenario.clientId !== undefined ? scenario.clientId : CLIENT_ID;
  sandbox.CONFIG.diagnostic = false;
  if (scenario.useBodyFallback !== undefined) {
    sandbox.CONFIG.useBodyFallback = scenario.useBodyFallback;
  }
  if (scenario.graphTimeoutMs) sandbox.CONFIG.graphTimeoutMs = scenario.graphTimeoutMs;

  return new Promise((resolve) => {
    sandbox.onMessageSendHandler({ completed: resolve });
  });
}

const QUOTED_MULTI = `Thanks.

From: Jane Doe <jane@vendor.com>
Sent: Monday, September 14, 2026 9:12 AM
To: Aaron H <aaron@bex.com>; Bob Smith <bob@bex.com>
Cc: Carol White <carol@vendor.com>
Subject: Re: Q3 invoice

Original.`;

const NEWEST = {
  isDraft: false,
  receivedDateTime: "2026-09-14T09:12:00Z",
  from: G("Jane Doe", "jane@vendor.com"),
  toRecipients: [G("Aaron H", "aaron@bex.com"), G("Bob Smith", "bob@bex.com")],
  ccRecipients: [G("Carol White", "carol@vendor.com")],
};

const OLDER = {
  isDraft: false,
  receivedDateTime: "2026-09-10T08:00:00Z",
  from: G("Someone Old", "old@vendor.com"),
  toRecipients: [G("Aaron H", "aaron@bex.com")],
  ccRecipients: [],
};

const THREAD = [NEWEST, OLDER];
const JANE = [R("Jane Doe", "jane@vendor.com")];

const tests = [
  { name: "Graph: reply drops Bob + Carol -> WARN",
    s: { composeType: "reply", to: JANE, graphMessages: THREAD },
    warn: true, expect: ["bob", "carol"] },

  { name: "Graph wins over a body that also has a quoted block",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, graphMessages: THREAD },
    warn: true, expect: ["bob", "carol"] },

  { name: "Graph: body EMPTY (new Outlook threading) -> still WARN",
    s: { composeType: "reply", to: JANE, body: "", graphMessages: THREAD },
    warn: true, expect: ["bob"] },

  { name: "Graph: reply all keeps everyone -> allow",
    s: { composeType: "replyAll",
         to: [R("Jane Doe", "jane@vendor.com"), R("Bob Smith", "bob@bex.com")],
         cc: [R("Carol White", "carol@vendor.com")],
         graphMessages: THREAD },
    warn: false },

  { name: "Graph: picks the newest message in the thread",
    s: { composeType: "reply", to: JANE, graphMessages: [OLDER, NEWEST] },
    warn: true, expect: ["bob", "carol"] },

  { name: "Graph: drafts in the thread are ignored",
    s: { composeType: "reply", to: JANE,
         graphMessages: [
           { isDraft: true, receivedDateTime: "2026-09-20T00:00:00Z",
             toRecipients: [G("Nobody", "nobody@x.com")] },
           NEWEST
         ] },
    warn: true, expect: ["bob", "carol"] },

  { name: "Token fails -> falls back to body parser",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, tokenFails: true },
    warn: true, expect: ["bob", "carol"] },

  { name: "MSAL missing (JS-only runtime) -> falls back to body parser",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, msal: false },
    warn: true, expect: ["bob"] },

  { name: "Graph 403 -> falls back to body parser",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, graphStatus: 403 },
    warn: true, expect: ["bob"] },

  { name: "Graph hangs -> times out, falls back, still warns",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, graphHangs: true, graphTimeoutMs: 150 },
    warn: true, expect: ["bob"] },

  { name: "Graph fails and fallback disabled -> allow, never blocks mail",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, graphFails: true, useBodyFallback: false },
    warn: false },

  { name: "clientId not configured yet -> falls back to body parser",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, clientId: "PASTE_AZURE_CLIENT_ID_HERE" },
    warn: true, expect: ["bob"] },

  { name: "Graph: original went only to me -> allow",
    s: { composeType: "reply", to: JANE,
         graphMessages: [{ isDraft: false, receivedDateTime: "2026-09-14T09:12:00Z",
                           toRecipients: [G("Aaron H", "aaron@bex.com")], ccRecipients: [] }] },
    warn: false },

  { name: "Graph: only other recipient is noreply@ -> allow",
    s: { composeType: "reply", to: JANE,
         graphMessages: [{ isDraft: false, receivedDateTime: "2026-09-14T09:12:00Z",
                           toRecipients: [G("Aaron H", "aaron@bex.com"), G("No Reply", "noreply@system.com")],
                           ccRecipients: [] }] },
    warn: false },

  { name: "No conversationId -> falls back to body parser",
    s: { composeType: "reply", to: JANE, body: QUOTED_MULTI, conversationId: null },
    warn: true, expect: ["bob"] },

  { name: "New mail -> allow",
    s: { composeType: "newMail", to: [R("Bob", "bob@bex.com")], body: "Hi", graphMessages: [] },
    warn: false },
];

(async () => {
  let pass = 0, fail = 0;

  for (const t of tests) {
    let r;
    try {
      r = await run(t.s);
    } catch (e) {
      r = { threw: String(e) };
    }

    const warned = r && r.allowEvent === false;
    let ok = warned === t.warn;

    if (ok && t.warn && t.expect) {
      const msg = (r.errorMessage || "").toLowerCase();
      ok = t.expect.every((n) => msg.includes(n));
    }

    if (ok) { pass++; console.log("  PASS  " + t.name); }
    else {
      fail++;
      console.log("  FAIL  " + t.name);
      console.log("        got: " + JSON.stringify(r));
    }
  }

  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
