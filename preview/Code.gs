/**
 * The Nest Log — backend.
 * Lives in the Apps Script of the Google Sheet "Nest Log data" in Anirban's Drive.
 * Deploy: Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 * Script properties (Project Settings > Script properties): PASSCODE (shared by both), OWNER_PIN (Anirban), BINDU_PIN (Bindu).
 * Sign-in emails live in the State tab, key "signin": {"emails":{"<email>":"Anirban","<email>":"Bindu"}}.
 * They are never sent to the app and are not kept in this repo. With no "signin" row, the passcode alone opens the app.
 */

const TABS = {
  Entries: ["id", "date", "who", "mood", "fight", "trigger", "felt", "needed", "change", "pauseUsed",
            "appreciation", "other", "context", "contextApp", "shotFile", "createdAt"],
  Feedback: ["id", "who", "text", "status", "reply", "createdAt"],
  Journal: ["id", "who", "date", "text", "createdAt"],
  State: ["key", "json", "updatedAt"]
};
const LIST_KEYS = ["todos", "shopping"];
const RULE_CHECKS = 3; // "I understand it", "I agree", "I'll follow it" — each of us ticks all three
const HIDDEN_STATE = ["accounts", "signin"]; // never sent to the app
const WHO = ["Bindu", "Anirban"];
const SHOT_FOLDER_ID = "15mOeSSfHazt_4G7Qx49N_aXTPy9XZumH"; // Drive folder "Nest Log screenshots"
const STARTER_RULES = [
  "We don't call each other names or insult each other, even when angry.",
  "We stick to today's issue and don't bring up old fights.",
  "We don't argue over text or call. Serious talks happen face to face.",
  "Either of us can say the pause word. We take 20–30 minutes apart, then the one who paused restarts the talk.",
  "We don't go silent on each other for more than a few hours without saying when we'll talk.",
  "When we talk it through, each person speaks without interruption and the other repeats back what they heard."
];
const TEXT_FIELDS = ["trigger", "felt", "needed", "change", "appreciation", "other"];

/** Run once from the editor: creates the tabs and default state. Safe to re-run. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(TABS[name]); sh.setFrozenRows(1); }
  });
  ss.getSheetByName("Entries").getRange("B:B").setNumberFormat("@"); // keep dates as plain YYYY-MM-DD text
  ss.getSheets().forEach(sh => { if (!TABS[sh.getName()] && sh.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(sh); });
  if (!getState_("settings")) setState_("settings", { aiPaused: false, nightlyPaused: false, remindersPaused: false });
  WHO.forEach(w => { if (!getState_("logstatus_" + w)) setState_("logstatus_" + w, { last: "", remind: true }); });
  if (!getState_("rules")) setState_("rules", { items: STARTER_RULES.map((t, i) => ({ id: "s" + (i + 1), text: t, source: "starter", added: "2026-10-25" })) });
}

function doGet() { return json_({ ok: true, app: "nest-log" }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({ error: "bad_request" }); }
  const props = PropertiesService.getScriptProperties();
  if (!req.pass || req.pass !== props.getProperty("PASSCODE")) return json_({ error: "wrong_passcode" });
  // Sign-in: email + shared passcode (approved by Anirban, 27 Oct 2026). Only the two emails in State "signin" are let in.
  let me = "";
  const signin = getState_("signin");
  if (signin && signin.emails) {
    me = signin.emails[String(req.email || "").trim().toLowerCase()] || "";
    if (WHO.indexOf(me) < 0) return json_({ error: "wrong_email" });
  }
  // Pause switches: open to whoever is signed in with their email (approved by Anirban, 30 Oct 2026).
  // OWNER_PIN / BINDU_PIN only matter as a fallback when there is no "signin" row.
  const isOwner = WHO.indexOf(me) >= 0 ||
    (!!req.ownerPin && (req.ownerPin === props.getProperty("OWNER_PIN") || req.ownerPin === props.getProperty("BINDU_PIN")));
  const d = req.data || {};
  // The signed-in person always writes, edits and journals as themselves; the app's "who" is ignored.
  if (WHO.indexOf(me) >= 0) d.who = me;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.action) {
      case "load": return json_(load_(isOwner, me));
      case "addEntry": return json_(addEntry_(d, props));
      case "updateEntry": return json_(updateEntry_(d));
      case "setRules": setState_("rules", { items: cleanRules_(d.items) }); return json_({ ok: true });
      case "addFeedback": return json_(addFeedback_(d));
      case "setRemind":
        if (WHO.indexOf(d.who) < 0) return json_({ error: "bad_request" });
        setState_("logstatus_" + d.who, Object.assign({}, getState_("logstatus_" + d.who) || {}, { remind: !!d.remind }));
        return json_({ ok: true });
      case "dismissSuggestion": return json_(dismiss_(d));
      case "saveList": return json_(saveList_(d));
      case "listOp": return json_(listOp_(d, me));
      case "proposeRule": return json_(proposeRule_(d, me));
      case "ruleTick": return json_(ruleTick_(d, me));
      case "dropRule": return json_(dropRule_(d));
      case "setJournalPass": return json_(setJournalPass_(d));
      case "journalLoad": return json_(journalLoad_(d));
      case "journalAdd": return json_(journalAdd_(d));
      case "setSetting":
        if (!isOwner) return json_({ error: "owner_only" });
        if (["aiPaused", "nightlyPaused", "remindersPaused"].indexOf(d.key) < 0) return json_({ error: "bad_request" });
        const s = getState_("settings") || {}; s[d.key] = !!d.value; setState_("settings", s);
        return json_({ ok: true });
      case "decide":
        // Approve / Reject is open to anyone signed in with the passcode (approved by Anirban, 27 Oct 2026).
        // The pause switches above need a signed-in email (or an owner PIN when there is no "signin" row).
        if (["approved", "rejected"].indexOf(d.status) < 0) return json_({ error: "bad_request" });
        const p = getState_("proposal");
        if (!p || p.status !== "pending") return json_({ error: "nothing_pending" });
        p.status = d.status; p.decidedAt = new Date().toISOString(); setState_("proposal", p);
        return json_({ ok: true });
      default: return json_({ error: "unknown_action" });
    }
  } finally { lock.releaseLock(); }
}

function load_(isOwner, me) {
  const state = {};
  rows_("State").forEach(r => { if (HIDDEN_STATE.indexOf(r.key) >= 0) return; try { state[r.key] = JSON.parse(r.json); } catch (x) {} }); // never send password digests or sign-in emails to the app
  const entries = rows_("Entries").slice(-300).reverse().map(r => { const o = Object.assign({}, r); delete o.shotFile; o.shotPending = !!r.shotFile; return o; });
  return { entries: entries, feedback: rows_("Feedback").slice(-200).reverse(), state: state, isOwner: isOwner, me: me };
}

function addEntry_(d, props) {
  if (WHO.indexOf(d.who) < 0) return { error: "bad_request" };
  const mood = Number(d.mood);
  if (!(mood >= 1 && mood <= 5)) return { error: "bad_request" };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(d.date) ? d.date : Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  const e = { id: Utilities.getUuid(), date: date, who: d.who, mood: mood,
              fight: ["none", "small", "big"].indexOf(d.fight) >= 0 ? d.fight : "none",
              pauseUsed: ["yes", "no", "na", ""].indexOf(d.pauseUsed) >= 0 ? d.pauseUsed : "",
              context: "", contextApp: "", shotFile: "", createdAt: new Date().toISOString() };
  TEXT_FIELDS.forEach(f => e[f] = String(d[f] || "").slice(0, 4000));
  if (d.shot && d.shot.base64) {
    const blob = Utilities.newBlob(Utilities.base64Decode(d.shot.base64), "image/jpeg", date + "-" + d.who + "-" + e.id.slice(0, 8) + ".jpg");
    e.shotFile = DriveApp.getFolderById(SHOT_FOLDER_ID).createFile(blob).getId();
  }
  SpreadsheetApp.getActive().getSheetByName("Entries").appendRow(TABS.Entries.map(k => e[k]));
  // "last" only moves forward, so writing yesterday's page today never undoes today's.
  const ls = Object.assign({ remind: true }, getState_("logstatus_" + d.who) || {});
  if (!ls.last || date > ls.last) ls.last = date;
  setState_("logstatus_" + d.who, ls);
  return { ok: true, id: e.id };
}

// Edits one Entries row. Only the writer named in d.who can edit a row they wrote.
// When signed in by email, doPost() sets d.who to the signed-in person, so nobody can edit the other's page.
function updateEntry_(d) {
  if (WHO.indexOf(d.who) < 0) return { error: "bad_request" };
  const sh = SpreadsheetApp.getActive().getSheetByName("Entries");
  const v = sh.getDataRange().getValues();
  const h = v[0];
  const idCol = h.indexOf("id"), whoCol = h.indexOf("who");
  for (let i = 1; i < v.length; i++) {
    if (v[i][idCol] !== d.id) continue;
    if (v[i][whoCol] !== d.who) return { error: "not_yours" };
    const row = h.map((k, j) => v[i][j]); // keeps id, date, createdAt and shotFile as they are
    const mood = Number(d.mood);
    if (mood >= 1 && mood <= 5) row[h.indexOf("mood")] = mood;
    if (["none", "small", "big"].indexOf(d.fight) >= 0) row[h.indexOf("fight")] = d.fight;
    if (["yes", "no", "na", ""].indexOf(d.pauseUsed) >= 0) row[h.indexOf("pauseUsed")] = d.pauseUsed;
    TEXT_FIELDS.forEach(f => {
      if (h.indexOf(f) >= 0 && d[f] !== undefined) row[h.indexOf(f)] = String(d[f]).slice(0, 4000);
    });
    sh.getRange(i + 1, 1, 1, h.length).setValues([row]);
    return { ok: true };
  }
  return { error: "not_found" };
}

function addFeedback_(d) {
  if (WHO.indexOf(d.who) < 0) return { error: "bad_request" };
  const t = String(d.text || "").trim().slice(0, 2000);
  if (!t) return { error: "bad_request" };
  const f = { id: Utilities.getUuid(), who: d.who, text: t, status: "new", reply: "", createdAt: new Date().toISOString() };
  SpreadsheetApp.getActive().getSheetByName("Feedback").appendRow(TABS.Feedback.map(k => f[k]));
  return { ok: true };
}

function dismiss_(d) {
  const s = getState_("suggestions");
  if (!s) return { ok: true };
  s.dismissed = (s.dismissed || []).concat([String(d.key || "")]).slice(-50);
  setState_("suggestions", s);
  return { ok: true };
}

// Shared to-do and shopping lists. Both of us can edit; stored in State.
function saveList_(d) {
  if (LIST_KEYS.indexOf(d.key) < 0) return { error: "bad_request" };
  const items = (Array.isArray(d.items) ? d.items : []).slice(0, 200).map(x => ({
    id: String(x.id || Utilities.getUuid()).slice(0, 40), text: String(x.text || "").slice(0, 300),
    done: !!x.done, by: WHO.indexOf(x.by) >= 0 ? x.by : ""
  })).filter(x => x.text);
  setState_(d.key, { items: items });
  return { ok: true };
}

// One change at a time, applied to the latest saved list, so two phones never overwrite each other.
function listOp_(d, me) {
  if (LIST_KEYS.indexOf(d.key) < 0) return { error: "bad_request" };
  let items = ((getState_(d.key) || {}).items || []).slice();
  if (d.op === "add") {
    const t = String(d.text || "").trim().slice(0, 300);
    if (!t) return { error: "bad_request" };
    const by = WHO.indexOf(me) >= 0 ? me : (WHO.indexOf(d.by) >= 0 ? d.by : "");
    if (items.length < 200) items.push({ id: String(d.id || Utilities.getUuid()).slice(0, 40), text: t, done: false, by: by });
  } else if (d.op === "toggle") {
    items = items.map(x => x.id === d.id ? Object.assign({}, x, { done: !!d.done }) : x);
  } else if (d.op === "remove") {
    items = items.filter(x => x.id !== d.id);
  } else return { error: "bad_request" };
  setState_(d.key, { items: items });
  return { ok: true, items: items };
}

// New rules wait in "pendingRules" until both of us have ticked all three checks.
function proposeRule_(d, me) {
  const t = String(d.text || "").trim().slice(0, 300);
  if (!t) return { error: "bad_request" };
  const pend = (getState_("pendingRules") || {}).items || [];
  if (pend.length >= 30) return { error: "too_many" };
  const ticks = {}; WHO.forEach(w => ticks[w] = [false, false, false]);
  pend.push({ id: Utilities.getUuid().slice(0, 12), text: t, source: d.source === "suggested" ? "suggested" : "manual",
              replaces: String(d.replaces || "").slice(0, 40), by: WHO.indexOf(me) >= 0 ? me : (WHO.indexOf(d.who) >= 0 ? d.who : ""),
              added: Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd"), ticks: ticks });
  setState_("pendingRules", { items: pend });
  return { ok: true, pending: pend };
}
function ruleTick_(d, me) {
  const who = WHO.indexOf(me) >= 0 ? me : d.who;
  const i = Number(d.i);
  if (WHO.indexOf(who) < 0 || !(i >= 0 && i < RULE_CHECKS)) return { error: "bad_request" };
  let pend = (getState_("pendingRules") || {}).items || [];
  const it = pend.filter(x => x.id === d.id)[0];
  if (!it) return { error: "not_found" };
  it.ticks = it.ticks || {};
  it.ticks[who] = (it.ticks[who] || [false, false, false]).slice(0, RULE_CHECKS);
  it.ticks[who][i] = !!d.value;
  let rules = (getState_("rules") || {}).items || [];
  const agreed = WHO.every(w => (it.ticks[w] || []).filter(Boolean).length === RULE_CHECKS);
  if (agreed) {
    const r = { id: it.replaces || Utilities.getUuid().slice(0, 12), text: it.text, source: it.source, added: it.added };
    if (it.replaces && rules.some(q => q.id === it.replaces)) rules = rules.map(q => q.id === it.replaces ? r : q);
    else rules.push(r);
    rules = cleanRules_(rules);
    setState_("rules", { items: rules });
    pend = pend.filter(x => x.id !== it.id);
  }
  setState_("pendingRules", { items: pend });
  return { ok: true, agreed: agreed, pending: pend, rules: rules };
}
function dropRule_(d) {
  const pend = ((getState_("pendingRules") || {}).items || []).filter(x => x.id !== d.id);
  setState_("pendingRules", { items: pend });
  return { ok: true, pending: pend };
}

// Private journal. Each of us sets a password; the journal is only returned with that password.
// Passwords are stored as salted SHA-256 digests in State key "accounts", never in plain text.
function digest_(salt, pw) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + "|" + pw));
}
function checkJournalPass_(who, pw) {
  if (WHO.indexOf(who) < 0 || !pw || String(pw).length < 6) return false;
  const acc = (getState_("accounts") || {})[who];
  return !!acc && acc.hash === digest_(acc.salt, String(pw));
}
function setJournalPass_(d) {
  if (WHO.indexOf(d.who) < 0) return { error: "bad_request" };
  const pw = String(d.password || "");
  if (pw.length < 6) return { error: "password_too_short" };
  const accounts = getState_("accounts") || {};
  const existing = accounts[d.who];
  if (existing && !checkJournalPass_(d.who, d.oldPassword)) return { error: "wrong_password" };
  const salt = Utilities.getUuid();
  accounts[d.who] = { salt: salt, hash: digest_(salt, pw) };
  setState_("accounts", accounts);
  return { ok: true };
}
function journalLoad_(d) {
  if (!checkJournalPass_(d.who, d.password)) return { error: "wrong_password" };
  const mine = rows_("Journal").filter(r => r.who === d.who).slice(-300).reverse();
  return { ok: true, entries: mine };
}
function journalAdd_(d) {
  if (!checkJournalPass_(d.who, d.password)) return { error: "wrong_password" };
  const t = String(d.text || "").trim().slice(0, 8000);
  if (!t) return { error: "bad_request" };
  const e = { id: Utilities.getUuid(), who: d.who, date: Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd"), text: t, createdAt: new Date().toISOString() };
  SpreadsheetApp.getActive().getSheetByName("Journal").appendRow(TABS.Journal.map(k => e[k]));
  return { ok: true, id: e.id };
}

function cleanRules_(items) {
  return (Array.isArray(items) ? items : []).slice(0, 40).map(r => ({
    id: String(r.id || Utilities.getUuid()).slice(0, 40), text: String(r.text || "").slice(0, 300),
    source: ["starter", "manual", "suggested"].indexOf(r.source) >= 0 ? r.source : "manual", added: String(r.added || "").slice(0, 10)
  })).filter(r => r.text);
}

// ---- helpers ----
function rows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const h = v[0];
  return v.slice(1).filter(r => r[0] !== "").map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]); return o; });
}
function getState_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName("State");
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) { try { return JSON.parse(v[i][1]); } catch (x) { return null; } }
  return null;
}
function setState_(key, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName("State");
  const v = sh.getDataRange().getValues();
  const row = [key, JSON.stringify(obj), new Date().toISOString()];
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) { sh.getRange(i + 1, 1, 1, 3).setValues([row]); return; }
  sh.appendRow(row);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
