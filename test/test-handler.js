// Harness: runs commands.js against a stubbed Office API.
const fs = require("fs");
const vm = require("vm");
const path = require("path");

const SRC = path.join(__dirname, "..", "commands.js");
const code = fs.readFileSync(process.argv[2] || SRC, "utf8");

const ME = { email: "aaron@bex.com", name: "Aaron H" };

function makeOffice(scenario) {
  const ok = (value) => (cb) => cb({ status: "succeeded", value });
  return {
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Text: "text" },
    onReady: () => {},
    actions: { associate: () => {} },
    context: {
      mailbox: {
        userProfile: { emailAddress: ME.email, displayName: ME.name },
        item: {
          getComposeTypeAsync: ok({ composeType: scenario.composeType }),
          body: { getAsync: (type, cb) => cb({ status: "succeeded", value: scenario.body }) },
          to: { getAsync: ok(scenario.to || []) },
          cc: { getAsync: ok(scenario.cc || []) },
          bcc: { getAsync: ok(scenario.bcc || []) },
        },
      },
    },
  };
}

function run(scenario) {
  const sandbox = { Office: makeOffice(scenario), console };
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox);

  let result = null;
  sandbox.onMessageSendHandler({ completed: (r) => { result = r; } });
  return result;
}

const R = (name, email) => ({ displayName: name, emailAddress: email });

const QUOTED_MULTI = `Thanks, will do.

From: Jane Doe <jane@vendor.com>
Sent: Monday, September 14, 2026 9:12 AM
To: Aaron H <aaron@bex.com>; Bob Smith <bob@bex.com>
Cc: Carol White <carol@vendor.com>
Subject: Re: Q3 invoice

Original text here.`;

const QUOTED_SINGLE = `Sure.

From: Jane Doe <jane@vendor.com>
Sent: Monday, September 14, 2026 9:12 AM
To: Aaron H <aaron@bex.com>
Subject: Lunch?

Original text.`;

const QUOTED_NAMES_ONLY = `Ok.

From: Jane Doe
Sent: Monday, September 14, 2026 9:12 AM
To: Aaron H; Bob Smith; Dana Lee
Subject: Internal sync

Original text.`;

const QUOTED_GERMAN = `Passt.

Von: Jane Doe <jane@vendor.com>
Gesendet: Montag, 14. September 2026 09:12
An: Aaron H <aaron@bex.com>; Bob Smith <bob@bex.com>
Betreff: Angebot

Original.`;

const QUOTED_SEMICOLON_SUBJECT = `Noted.

From: Jane Doe <jane@vendor.com>
Sent: Monday, September 14, 2026 9:12 AM
To: Aaron H <aaron@bex.com>
Subject: Budget; forecast; and headcount

Original.`;

const QUOTED_NOREPLY = `Ok.

From: Jane Doe <jane@vendor.com>
Sent: Monday, September 14, 2026 9:12 AM
To: Aaron H <aaron@bex.com>; noreply@system.com
Subject: Ticket 88

Original.`;

const tests = [
  ["Reply, original had Bob + Carol -> WARN",
    { composeType: "reply", body: QUOTED_MULTI, to: [R("Jane Doe", "jane@vendor.com")] },
    true, ["bob", "carol"]],

  ["Reply All on same thread -> allow",
    { composeType: "replyAll", body: QUOTED_MULTI,
      to: [R("Jane Doe", "jane@vendor.com"), R("Bob Smith", "bob@bex.com")],
      cc: [R("Carol White", "carol@vendor.com")] },
    false],

  ["Reply, original was to me only -> allow",
    { composeType: "reply", body: QUOTED_SINGLE, to: [R("Jane Doe", "jane@vendor.com")] },
    true === false, []],

  ["Reply, display-names only (Exchange) -> WARN",
    { composeType: "reply", body: QUOTED_NAMES_ONLY, to: [R("Jane Doe", "") ] },
    true, ["bob smith", "dana lee"]],

  ["Reply, German client -> WARN",
    { composeType: "reply", body: QUOTED_GERMAN, to: [R("Jane Doe", "jane@vendor.com")] },
    true, ["bob"]],

  ["Reply, semicolons in Subject line -> allow (no false positive)",
    { composeType: "reply", body: QUOTED_SEMICOLON_SUBJECT, to: [R("Jane Doe", "jane@vendor.com")] },
    false],

  ["Reply, only other recipient is noreply@ -> allow",
    { composeType: "reply", body: QUOTED_NOREPLY, to: [R("Jane Doe", "jane@vendor.com")] },
    false],

  ["New mail, no quote -> allow",
    { composeType: "newMail", body: "Hi there", to: [R("Bob", "bob@bex.com")] },
    false],

  ["Reply, user already re-added Bob manually -> allow",
    { composeType: "reply", body: QUOTED_MULTI,
      to: [R("Jane Doe", "jane@vendor.com"), R("Bob Smith", "bob@bex.com"), R("Carol White", "carol@vendor.com")] },
    false],
];


const QUOTED_UNKNOWN_LOCALE = `Ok.

From: Jane Doe <jane@vendor.com>
Odoslane: pondelok, September 14, 2026
To: Aaron H <aaron@bex.com>
Subject: Test

Original.`;

tests.push(["Reply, unknown-locale date line -> allow (no false positive)",
  { composeType: "reply", body: QUOTED_UNKNOWN_LOCALE, to: [R("Jane Doe", "jane@vendor.com")] },
  false]);

// Fix the third test's expectation (single-recipient original should allow).
tests[2][2] = false;

let pass = 0, fail = 0;
for (const [name, scenario, shouldWarn, expectNames] of tests) {
  const r = run(scenario);
  const warned = r && r.allowEvent === false;
  let ok = warned === shouldWarn;

  if (ok && shouldWarn && expectNames) {
    const msg = (r.errorMessage || "").toLowerCase();
    ok = expectNames.every((n) => msg.includes(n));
  }

  if (ok) { pass++; console.log("  PASS  " + name); }
  else {
    fail++;
    console.log("  FAIL  " + name);
    console.log("        got: " + JSON.stringify(r));
  }
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
