/*
 * Reply-All Reminder for new Outlook / Outlook on the web
 * -------------------------------------------------------
 * Fires on Send. If you used Reply (not Reply All) on a message that had
 * other recipients, it warns you and lists who is about to be dropped.
 *
 * Everything runs locally in the add-in runtime. No network calls, no
 * account permissions beyond reading the message you are composing.
 */

/* ------------------------------------------------------------------ */
/* CONFIG - edit these                                                 */
/* ------------------------------------------------------------------ */

var CONFIG = {
  // Your own addresses. You are never counted as a "dropped" recipient.
  // Add every alias you receive mail at.
  myAddresses: [
    // "aaron@example.com",
  ],

  // Never warn about these addresses or domain fragments.
  ignore: [
    "noreply@",
    "no-reply@",
    "donotreply@"
  ],

  // Warn only when at least this many people would be left out.
  minDropped: 1,

  // Header labels used to find the To/Cc lines in the quoted original.
  // Detection still works without a label match - these just improve it.
  toLabels: ["to", "an", "a", "para", "aan", "til", "till", "kenelle", "do", "komu"],
  ccLabels: ["cc", "copy", "kopie", "copia", "kopia"],
  fromLabels: ["from", "von", "de", "da", "van", "fran", "fra", "lahettaja", "od"],
  sentLabels: ["sent", "date", "gesendet", "envoye", "enviado", "inviato", "verzonden", "skickat", "sendt", "lahetetty", "wyslano"],
  subjectLabels: ["subject", "betreff", "objet", "asunto", "oggetto", "onderwerp", "amne", "emne", "aihe", "temat"]
};

/* ------------------------------------------------------------------ */
/* Entry point                                                         */
/* ------------------------------------------------------------------ */

function onMessageSendHandler(event) {
  var composeType = null;
  var currentRecipients = [];
  var bodyText = "";

  function finish() {
    try {
      // Only interested in Reply. Reply All and new mail are fine, and a
      // deliberately trimmed Reply All is the user's own decision.
      if (composeType !== "reply") return allow(event);

      var block = extractHeaderBlock(bodyText);
      if (!block) return allow(event);

      var original = parseParticipants(block);
      if (!original.length) return allow(event);

      var dropped = original.filter(function (p) {
        return !isMe(p) && !isIgnored(p) && !isAlreadyIncluded(p, currentRecipients);
      });

      if (dropped.length < CONFIG.minDropped) return allow(event);

      event.completed({
        allowEvent: false,
        errorMessage: buildWarning(dropped)
      });
    } catch (e) {
      // Never let a bug in here block a send.
      allow(event);
    }
  }

  // Three independent lookups. Fire them together, join on a counter.
  var pending = 3;
  function step() { if (--pending === 0) finish(); }

  var item = Office.context.mailbox.item;

  item.getComposeTypeAsync(function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
      composeType = r.value.composeType;
    }
    step();
  });

  item.body.getAsync(Office.CoercionType.Text, function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      bodyText = r.value || "";
    }
    step();
  });

  getAllRecipients(item, function (list) {
    currentRecipients = list;
    step();
  });
}

function allow(event) {
  event.completed({ allowEvent: true });
}

/* ------------------------------------------------------------------ */
/* Current recipients                                                  */
/* ------------------------------------------------------------------ */

function getAllRecipients(item, callback) {
  var out = [];
  var pending = 3;
  function step() { if (--pending === 0) callback(out); }

  function collect(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
      result.value.forEach(function (r) {
        out.push({ email: norm(r.emailAddress), name: norm(r.displayName) });
      });
    }
    step();
  }

  item.to.getAsync(collect);
  item.cc.getAsync(collect);
  item.bcc.getAsync(collect);
}

/* ------------------------------------------------------------------ */
/* Quoted-original parsing                                             */
/* ------------------------------------------------------------------ */

/*
 * Pull out the header block of the most recent quoted message - the
 * From/Sent/To/Cc/Subject lines Outlook inserts above the quoted text.
 * Returns an array of lines, or null.
 */
function extractHeaderBlock(body) {
  var lines = body.split(/\r?\n/);

  for (var i = 0; i < lines.length; i++) {
    if (!matchesLabel(lines[i], CONFIG.fromLabels)) continue;

    var block = [];
    var end = Math.min(i + 10, lines.length);
    for (var j = i; j < end; j++) {
      var line = lines[j];
      // A blank line ends the block, but only once we have a couple of lines.
      if (!line.trim() && block.length >= 2) break;
      block.push(line);
      if (matchesLabel(line, CONFIG.subjectLabels)) break;
    }
    return block.length >= 2 ? block : null;
  }
  return null;
}

/*
 * Participants of the original message, excluding its sender (who is
 * already the To: of your reply).
 */
function parseParticipants(block) {
  var out = [];

  block.forEach(function (line, idx) {
    if (idx === 0) return;                                 // From line
    if (matchesLabel(line, CONFIG.subjectLabels)) return;  // Subject line
    if (matchesLabel(line, CONFIG.fromLabels)) return;
    if (matchesLabel(line, CONFIG.sentLabels)) return;     // Sent/Date line

    var isLabelled =
      matchesLabel(line, CONFIG.toLabels) || matchesLabel(line, CONFIG.ccLabels);

    var entries = splitEntries(stripLabel(line));

    // A date line in an unrecognised locale would otherwise read as two
    // name-like tokens. Nothing with a year or a clock time is a recipient.
    if (!isLabelled && looksLikeTimestamp(line)) return;

    // A labelled To/Cc line is trusted outright. An unlabelled line counts
    // only if it actually reads like a recipient list - that is what keeps
    // non-English clients working without any config.
    if (!isLabelled && !looksLikeRecipientList(entries)) return;

    entries.forEach(function (e) {
      var p = parseEntry(e);
      if (p) out.push(p);
    });
  });

  return dedupe(out);
}

function splitEntries(value) {
  // ';' always separates. ',' only separates outside an <...> address.
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

  // Display name only - common for internal Exchange recipients.
  if (entry.length >= 2 && entry.length <= 80 && entry.indexOf("@") === -1 && /[A-Za-z]/.test(entry)) {
    return { email: "", name: norm(entry.replace(/["']/g, "")) };
  }
  return null;
}

/* A 4-digit year or a clock time means this is a date line, not recipients. */
function looksLikeTimestamp(line) {
  return /\b\d{4}\b/.test(line) || /\b\d{1,2}[:.]\d{2}\b/.test(line);
}

/* Two or more entries that each read like a person or an address. */
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
  var profile = Office.context.mailbox.userProfile || {};
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
  return recipients.some(function (r) {
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

Office.onReady();

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
