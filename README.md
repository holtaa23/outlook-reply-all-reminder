# Reply-All Reminder

Warns you when you hit **Reply** on a message that had other recipients, and
names exactly who is about to be dropped from the thread.

This is the inverse of Outlook's built-in Reply All warning — it catches the
mistake of replying too narrowly, not too broadly.

Two implementations, same logic:

| | Classic Outlook (Windows) | New Outlook / Outlook on the web |
|---|---|---|
| Folder | `classic-outlook/` | repo root |
| Tech | VBA macro | Office Add-in (`OnMessageSend`) |
| Setup | 5 minutes, no hosting | needs HTTPS hosting + sideload |
| Cost | free | free |

---

## How it decides

1. Is this a **Reply**? (Not Reply All, not Forward, not a new message.)
   Reply All is left alone — if you trimmed recipients off a Reply All, that
   was deliberate.
2. Read the quoted original's `From/Sent/To/Cc/Subject` header block out of
   the message body.
3. Everyone on the original's To/Cc, minus you, minus the ignore list, minus
   anyone already on your reply → **the people who'd be dropped**.
4. If that list isn't empty, prompt.

No network calls, no mailbox-wide permissions, nothing leaves your machine.

**The tradeoff:** it reads the quoted header block rather than querying the
server for the original message. That keeps it free, offline, and dependency-free.
The cost is that it can't see anything if you've deleted the quoted text from
your reply. It handles English plus ~10 other locales out of the box, and
falls back to structural detection (a `;`-separated list that looks like
people, and isn't a date line) when the label language is unrecognised.

---

## Classic Outlook setup

1. Enable the Developer tab: **File → Options → Customize Ribbon** → tick
   **Developer**.
2. **Developer → Visual Basic** (or `Alt+F11`).
3. **File → Import File…** → pick `classic-outlook/modReplyAllReminder.bas`.
4. In the tree on the left, double-click **ThisOutlookSession**. Paste in the
   contents of `classic-outlook/ThisOutlookSession.txt`.
   (If you already have an `Application_ItemSend`, just add the one
   `CheckReplyAll` line to the sub you already have.)
5. **Debug → Compile Project** — should report nothing.
6. Save (`Ctrl+S`), then restart Outlook.

If the macro doesn't fire, it's macro security:
**File → Options → Trust Center → Trust Center Settings → Macro Settings**.
Either sign the project with a self-cert (`SelfCert.exe`, ships with Office)
and choose *Notifications for digitally signed macros only*, or pick
*Notifications for all macros* and approve it at each start.

**Config** — constants at the top of `modReplyAllReminder.bas`:

- `MY_EXTRA_ADDRESSES` — aliases you receive at, semicolon-separated. Your
  Outlook account addresses are detected automatically; add anything else.
- `IGNORE_FRAGMENTS` — never warn about these. Add ticketing systems and
  distribution lists that fan out on their own.
- `MIN_DROPPED` — set to `2` if single-person drops are too chatty.

---

## New Outlook / Outlook on the web setup

New Outlook runs no VBA and no COM add-ins, so this has to be a web add-in.
The files are static — any HTTPS host works.

### 1. Host the files

Host `commands.js`, `commands.html` and `assets/`
somewhere HTTPS. GitHub Pages is free and fine:

```bash
git init && git add . && git commit -m "Reply-All Reminder"
```

Push to a repo, then **Settings → Pages → Deploy from branch → root**. You'll
get `https://<user>.github.io/<repo>/`.

### 2. Point the manifest at it

Edit `manifest.xml` and replace every occurrence of
`https://REPLACE-ME.example.com/` with your base URL (keep the trailing slash).

```bash
sed -i 's#https://REPLACE-ME.example.com/#https://YOURUSER.github.io/YOURREPO/#g' manifest.xml
```

### 3. Set your addresses

In `commands.js`, add your aliases to `CONFIG.myAddresses`. Your primary
mailbox address is detected automatically, so this is only needed if you
receive mail at more than one address.

### 4. Sideload

In **Outlook on the web** (works for new Outlook for Windows too — they share
add-ins):

1. Open any message → **⋯ → Get Add-ins** (or the Apps icon → **Add-ins**).
2. **My add-ins → Custom Addins → Add a custom add-in → Add from File**.
3. Pick `manifest.xml`, accept the sideload warning.
4. Restart Outlook. Event-based add-ins need a restart to register.

Give it a few minutes on first install — Outlook caches the manifest.

### Verify it works

Send yourself a test: have a message addressed to you **and** one other
person, then hit Reply and send. You should get a "Send Anyway / Don't Send"
prompt naming the other person.

### If your tenant blocks sideloading

Corporate tenants often disable user-installed add-ins. Then it's either
**Microsoft 365 admin center → Settings → Integrated apps → Upload custom
apps** (needs an admin), or ask IT to deploy it. The manifest is the whole
ask — there's no backend to approve.

---

## Tests

The detection logic has a test harness that stubs the Office API, so you can
change the config and confirm you haven't broken anything:

```bash
node test/test-handler.js
```

Covers: reply with dropped recipients, Reply All, single-recipient originals,
display-name-only Exchange recipients, a non-English client, semicolons in the
subject line, ignore-list addresses, and manually re-added recipients.

---

## Known limits

- **Quoted text deleted** → nothing to read, no warning. Applies to both versions.
- **Plain-text-only threads with no header block** (some mailing lists strip it)
  → no warning.
- **New Outlook prompt text is capped** at roughly 150 characters by Microsoft,
  so it lists up to 3 names then `+N more`. The Classic version has no such cap
  and lists everyone.
- **Classic version can't tell Reply from Reply All directly** — it uses the
  In-Reply-To header to exclude forwards, then compares recipient sets. A Reply
  All with everyone intact produces no warning because nobody is missing, which
  is the correct outcome either way.
