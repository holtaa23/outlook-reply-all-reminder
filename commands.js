/*
 * Reply-All Reminder for new Outlook / Outlook on the web
 * -------------------------------------------------------
 * Fires on Send. If you used Reply on a message that had other recipients,
 * it warns you and names who is about to be dropped from the thread.
 *
 * It asks Microsoft Graph for the ORIGINAL message and reads its real To and
 * Cc lists. That is the whole point of this version: the previous one parsed
 * the quoted "From:/To:/Cc:" block out of the reply body, and new Outlook
 * threads replies instead of inlining that block, so there was often nothing
 * there to read.
 *
 * The body parser is kept only as a fallback for when Graph is unreachable
 * (offline, token expired, classic Outlook's JavaScript-only runtime where
 * MSAL cannot load). Graph always wins when both are available. Set
 * CONFIG.useBodyFallback to false to turn the fallback off entirely.
 *
 * SETUP: CONFIG.clientId below must hold the Application (client) ID of the
 * Azure app registration. Until it does, this falls back to body parsing.
 */

/* ------------------------------------------------------------------ */
/* CONFIG                                                              */
/* ------------------------------------------------------------------ */

var CONFIG = {
  // Application (client) ID from the Azure app registration.
  clientId: "PASTE_AZURE_CLIENT_ID_HERE",

  // Permission asked of Graph. Mail.Read is the narrowest that can read the
  // original message's recipients. Delegated - it can only ever see mail the
  // signed-in user can already see.
  graphScopes: ["Mail.Read"],

  // Give up on Graph after this long and fall back. The send is blocked while
  // this runs, so it has to stay short.
  graphTimeoutMs: 4000,

  // Read the quoted header block when Graph is unavailable.
  useBodyFallback: true,

  // TEMPORARY. Prompts on every reply with what the handler actually saw,
  // including which source the data came from. Set to false for normal use.
  diagnostic: true,

  // Extra addresses that count as "you". Your primary mailbox address is
  // detected automatically; add aliases here.
  myAddresses: [],

  // Never warn about these addresses or domain fragments.
  ignore: ["noreply@", "no-reply@", "donotreply@"],

  // Warn only when at least this many people would be left out.
  minDropped: 1,

  // Header labels, used only by the body-parsing fallback.
  toLabels: ["to", "an", "a", "para", "aan", "til", "till", "kenelle", "do", "komu"],
  ccLabels: ["cc", "copy", "kopie", "copia", "kopia"],
  fromLabels: ["from", "von", "de", "da", "van", "fran", "fra", "lahettaja", "od"],
  sentLabels: ["sent", "date", "gesendet", "envoye", "enviado", "inviato", "verzonden", "skickat", "sendt", "lahetetty", "wyslano"],
  subjectLabels: ["subject", "betreff", "objet", "asunto", "oggetto", "onderwerp", "amne", "emne", "aihe", "temat"]
};

var GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

/* Reused across sends within one runtime lifetime. The runtime is short-lived,
 * so this helps on a burst of replies rather than across a whole session. */
var msalInstance = null;

/* ------------------------------------------------------------------ */
/* Entry point                                                         */
/* ------------------------------------------------------------------ */

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var state = { source: "none", error: "" };

  // Office.onReady does not run for event handlers, so everything starts here.
  Promise.all([
    getComposeType(item),
    getAllRecipients(item),
    getBodyText(item)
  ])
    .then(function (results) {
      state.composeType = results[0];
      state.recipients = results[1];
      state.bodyText = results[2];

      return getOriginalParticipants(item, state);
    })
    .then(function (original) {
      state.original = original || [];

      state.dropped = state.original.filter(function (p) {
        return !isMe(p) && !isIgnored(p) && !isAlreadyIncluded(p, state.recipients);
      });

      decide(event, state);
    })
    .catch(function (e) {
      state.error = String((e && e.message) || e).slice(0, 80);
      if (CONFIG.diagnostic) {
        return event.completed({
          allowEvent: false,
          errorMessage: "DIAG threw: " + state.error
        });
      }
      // Never block a send because this add-in failed.
      allow(event);
    });
}

function decide(event, state) {
  if (CONFIG.diagnostic) {
    return event.completed({
      allowEvent: false,
      errorMessage: buildDiagnostic(state)
    });
  }

  // Only Reply is interesting. A trimmed Reply All is a deliberate choice.
  if (state.composeType !== "reply") return allow(event);
  if (state.dropped.length < CONFIG.minDropped) return allow(event);

  event.completed({
    allowEvent: false,
    errorMessage: buildWarning(state.dropped)
  });
}

function allow(event) {
  event.completed({ allowEvent: true });
}

/* ------------------------------------------------------------------ */
/* The original message's recipients - Graph first, body as fallback   */
/* ------------------------------------------------------------------ */

function getOriginalParticipants(item, state) {
  return fromGraph(item, state)
    .then(function (people) {
      if (people && people.length) {
        state.source = "graph";
        return people;
      }
      return fromBody(state);
    })
    .catch(function (e) {
      state.error = String((e && e.message) || e).slice(0, 60);
      return fromBody(state);
    });
}

function fromBody(state) {
  if (!CONFIG.useBodyFallback) return [];

  var block = extractHeaderBlock(state.bodyText || "");
  if (!block) return [];

  var people = parseParticipants(block);
  if (people.length) state.source = "body";
  return people;
}

/*
 * Ask Graph for the messages in this conversation and read the real To/Cc off
 * the most recent one that isn't the draft being composed.
 */
function fromGraph(item, state) {
  var conversationId = item.conversationId;

  if (!conversationId) return Promise.resolve([]);
  if (!CONFIG.clientId || CONFIG.clientId.indexOf("PASTE_") === 0) {
    state.error = "no clientId";
    return Promise.resolve([]);
  }

  return withTimeout(
    getGraphToken().then(function (token) {
      var url =
        GRAPH_ROOT +
        "/me/messages?$filter=conversationId eq '" +
        encodeURIComponent(conversationId) +
        "'&$select=from,toRecipients,ccRecipients,receivedDateTime,isDraft&$top=25";

      return fetch(url, {
        headers: { Authorization: "Bearer " + token }
      }).then(function (res) {
        if (!res.ok) throw new Error("graph " + res.status);
        return res.json();
      });
    }),
    CONFIG.graphTimeoutMs
  ).then(function (data) {
    var messages = (data && data.value) || [];

    var original = messages
      .filter(function (m) { return !m.isDraft && m.receivedDateTime; })
      .sort(function (a, b) {
        return new Date(b.receivedDateTime) - new Date(a.receivedDateTime);
      })[0];

    if (!original) return [];

    // The sender is already the To: of your reply, so only To and Cc matter.
    return dedupe(
      []
        .concat(original.toRecipients || [], original.ccRecipients || [])
        .map(function (r) {
          var a = (r && r.emailAddress) || {};
          return { email: norm(a.address), name: norm(a.name) };
        })
        .filter(function (p) { return p.email || p.name; })
    );
  });
}

function getGraphToken() {
  return initMsal().then(function (instance) {
    var accounts = instance.getAllAccounts();
    var request = { scopes: CONFIG.graphScopes };
    if (accounts && accounts.length) request.account = accounts[0];

    // Silent only. An interactive popup in the middle of a send would be
    // hostile, and may not be permitted from the event runtime at all.
    return instance.acquireTokenSilent(request).then(function (result) {
      return result.accessToken;
    });
  });
}

function initMsal() {
  if (msalInstance) return Promise.resolve(msalInstance);

  // MSAL is loaded as a global by commands.html. In classic Outlook's
  // JavaScript-only runtime there is no page, so this will be undefined and
  // we fall back to body parsing.
  if (typeof msal === "undefined" || !msal.createNestablePublicClientApplication) {
    return Promise.reject(new Error("msal unavailable"));
  }

  return msal
    .createNestablePublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: "https://login.microsoftonline.com/common"
      },
      cache: { cacheLocation: "localStorage" }
    })
    .then(function (instance) {
      msalInstance = instance;
      return instance;
    });
}

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise(function (_, reject) {
      setTimeout(function () { reject(new Error("timeout")); }, ms);
    })
  ]);
}

/* ------------------------------------------------------------------ */
/* Office.js wrappers                                                  */
/* ------------------------------------------------------------------ */

function getComposeType(item) {
  return new Promise(function (resolve) {
    if (!item.getComposeTypeAsync) return resolve(null);
    item.getComposeTypeAsync(function (r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded && r.value
        ? r.value.composeType
        : null);
    });
  });
}

function getBodyText(item) {
  return new Promise(function (resolve) {
    item.body.getAsync(Office.CoercionType.Text, function (r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value || "" : "");
    });
  });
}

function getAllRecipients(item) {
  function read(field) {
    return new Promise(function (resolve) {
      field.getAsync(function (r) {
        resolve(r.status === Office.AsyncResultStatus.Succeeded && r.value ? r.value : []);
      });
    });
  }

  return Promise.all([read(item.to), read(item.cc), read(item.bcc)]).then(function (lists) {
    return [].concat(lists[0], lists[1], lists[2]).map(function (r) {
      return { email: norm(r.emailAddress), name: norm(r.displayName) };
    });
  });
}

/* ------------------------------------------------------------------ */
/* Body-parsing fallback                                               */
/* ------------------------------------------------------------------ */

function extractHeaderBlock(body) {
  var lines = body.split(/\r?\n/);
  var block;

  for (var i = 0; i < lines.length; i++) {
    if (!matchesLabel(lines[i], CONFIG.fromLabels)) continue;

    block = [];
    var end = Math.min(i + 10, lines.length);
    for (var j = i; j < end; j++) {
      if (!lines[j].trim() && block.length >= 2) break;
      block.push(lines[j]);
      if (matchesLabel(lines[j], CONFIG.subjectLabels)) break;
    }
    return block.length >= 2 ? block : null;
  }
  return null;
}

function parseParticipants(block) {
  var out = [];

  block.forEach(function (line, idx) {
    if (idx === 0) return;
    if (matchesLabel(line, CONFIG.subjectLabels)) return;
    if (matchesLabel(line, CONFIG.fromLabels)) return;
    if (matchesLabel(line, CONFIG.sentLabels)) return;

    var labelled =
      matchesLabel(line, CONFIG.toLabels) || matchesLabel(line, CONFIG.ccLabels);
    var entries = splitEntries(stripLabel(line));

    if (!labelled && looksLikeTimestamp(line)) return;
    if (!labelled && !looksLikeRecipientList(entries)) return;

    entries.forEach(function (e) {
      var p = parseEntry(e);
      if (p) out.push(p);
    });
  });

  return dedupe(out);
}

function splitEntries(value) {
  return value
    .split(/;|,(?![^<]*>)/)
    .map(function (s) { return s.trim(); })
    .filter(Boolean);
}

function parseEntry(entry) {
  var angle = entry.match(/<([^>]+@[^>]+)>/);
  if (angle) {
    return {
      email: norm(angle[1]),
      name: norm(entry.slice(0, entry.indexOf("<")).replace(/["']/g, ""))
    };
  }

  var bare = entry.match(/[^\s<>()[\],;:]+@[^\s<>()[\],;:]+\.[A-Za-z]{2,}/);
  if (bare) return { email: norm(bare[0]), name: "" };

  if (entry.length >= 2 && entry.length <= 80 && entry.indexOf("@") === -1 && /[A-Za-z]/.test(entry)) {
    return { email: "", name: norm(entry.replace(/["']/g, "")) };
  }
  return null;
}

function looksLikeTimestamp(line) {
  return /\b\d{4}\b/.test(line) || /\b\d{1,2}[:.]\d{2}\b/.test(line);
}

function looksLikeRecipientList(entries) {
  if (entries.length < 2) return false;
  var plausible = entries.filter(function (e) {
    var t = e.trim();
    return t.indexOf("@") !== -1 || /^[A-Za-z'.-]+(\s+[A-Za-z'.-]+){0,3}$/.test(t);
  });
  return plausible.length >= 2;
}

/* ------------------------------------------------------------------ */
/* Matching                                                            */
/* ------------------------------------------------------------------ */

function isMe(p) {
  var mine = CONFIG.myAddresses.map(norm);
  var profile = (Office.context.mailbox && Office.context.mailbox.userProfile) || {};
  var box = norm(profile.emailAddress);
  var display = norm(profile.displayName);

  if (p.email && (mine.indexOf(p.email) !== -1 || p.email === box)) return true;
  if (!p.email && p.name && p.name === display) return true;
  return false;
}

function isIgnored(p) {
  if (!p.email) return false;
  return CONFIG.ignore.some(function (frag) {
    return p.email.indexOf(norm(frag)) !== -1;
  });
}

function isAlreadyIncluded(p, recipients) {
  return (recipients || []).some(function (r) {
    if (p.email && r.email) return p.email === r.email;
    if (p.name && r.name) return p.name === r.name;
    return false;
  });
}

function dedupe(list) {
  var seen = {};
  return list.filter(function (p) {
    var key = p.email || p.name;
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function norm(s) {
  return (s || "").trim().toLowerCase();
}

function labelOf(line) {
  var m = line.match(/^\s*\*?\s*([^:]{1,24}):/);
  return m ? norm(m[1]) : null;
}

function matchesLabel(line, labels) {
  var label = labelOf(line);
  if (label === null) return false;
  return labels.some(function (l) { return label === norm(l); });
}

function stripLabel(line) {
  var i = line.indexOf(":");
  return i >= 0 && labelOf(line) !== null ? line.slice(i + 1).trim() : line.trim();
}

function buildDiagnostic(state) {
  var names = (state.dropped || [])
    .map(function (p) { return p.name || p.email; })
    .filter(Boolean)
    .join(",");

  var msg =
    "DIAG src=" + state.source +
    " ct=" + (state.composeType || "?") +
    " rcp=" + ((state.recipients || []).length) +
    " orig=" + ((state.original || []).length) +
    " drop=" + ((state.dropped || []).length) +
    (state.error ? " err=" + state.error : "") +
    (names ? " [" + names + "]" : "");

  return msg.length > 140 ? msg.slice(0, 137) + "..." : msg;
}

function buildWarning(dropped) {
  var names = dropped
    .map(function (p) { return p.name || p.email; })
    .filter(Boolean);

  var shown = names.slice(0, 3).join(", ");
  var extra = names.length > 3 ? " +" + (names.length - 3) + " more" : "";
  var who = shown ? shown + extra : dropped.length + " other recipient(s)";

  return "You clicked Reply, not Reply All. Left off: " + who + ". Send anyway?";
}

/* ------------------------------------------------------------------ */
/* Registration                                                        */
/* ------------------------------------------------------------------ */

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
