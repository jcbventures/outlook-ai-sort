"use strict";

/* ══════════════════════════════════════════════════════════════
   AI Mapadvies – taskpane.js
   Modes:
     • Single  – analyseer huidige e-mail, stel map voor
     • Bulk    – loop door inbox, mail voor mail:
                 ✅ Actie (vlag) | ⏭ Overslaan | 📁 Verplaatsen
   ══════════════════════════════════════════════════════════════ */

// ── Shared state ──────────────────────────────────────────────
let allFolders        = [];     // { id, changeKey, name }
let singleFolder      = null;   // selected folder in single mode
let bulkEmails        = [];     // inbox items
let bulkIndex         = 0;
let bulkAiResult      = null;   // AI result for current bulk email
let bulkSelectedFolder= null;
let bulkStats         = { actie: 0, overgeslagen: 0, verplaatst: 0 };

// ── Office ready ──────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  if (!localStorage.getItem("claudeApiKey")) {
    showScreen("settings");
  } else {
    showScreen("home");
  }
});

// ── Screen navigation ─────────────────────────────────────────
function showScreen(name) {
  document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
  const el = document.getElementById("screen-" + name);
  if (el) el.classList.add("active");

  if (name === "settings") {
    document.getElementById("inp-key").value = localStorage.getItem("claudeApiKey") || "";
    document.getElementById("inp-ctx").value = localStorage.getItem("extraContext") || "";
  }
}

// ── Settings ──────────────────────────────────────────────────
function saveSettings() {
  const key = document.getElementById("inp-key").value.trim();
  if (!key.startsWith("sk-ant-")) {
    setEl("settings-status", '<div class="status status-err">Voer een geldige API-sleutel in (begint met sk-ant-).</div>');
    return;
  }
  localStorage.setItem("claudeApiKey", key);
  localStorage.setItem("extraContext", document.getElementById("inp-ctx").value.trim());
  setEl("settings-status", '<div class="status status-ok">✅ Opgeslagen!</div>');
  setTimeout(() => showScreen("home"), 800);
}

// ══════════════════════════════════════════════════════════════
//  SINGLE MODE
// ══════════════════════════════════════════════════════════════
async function startSingleMode() {
  showScreen("single");
  show("single-loading"); hide("single-result"); hide("single-error");
  singleFolder = null;

  try {
    const item    = Office.context.mailbox.item;
    const subject = item.subject || "(geen onderwerp)";
    const from    = item.from ? (item.from.displayName || item.from.emailAddress) : "onbekend";

    const [folders, body] = await Promise.all([getFolders(), getBody(item)]);
    allFolders = folders;

    const ai = await callClaude(subject, from, body, folders);
    singleFolder = folders.find(f => matchFolder(f, ai.suggested_folder)) || null;

    setEl("s-from",    esc(from));
    setEl("s-subject", esc(subject));
    setBadge("s-badge",   ai.action_type);
    setEl("s-summary", esc(ai.action_summary));
    setEl("s-folder",  esc(ai.suggested_folder));
    setConfBar("s-conf-bar", "s-conf-lbl", ai.confidence);
    setEl("s-reason",  esc(ai.folder_reason));
    renderFolderList("s-list", folders, singleFolder, (f) => { singleFolder = f; });

    hide("single-loading"); show("single-result");
  } catch (e) {
    console.error(e);
    setEl("s-err-msg", "❌ " + e.message);
    hide("single-loading"); show("single-error");
  }
}

async function moveEmailSingle() {
  if (!singleFolder) { showStatus("s-status", "Selecteer eerst een map.", "err"); return; }
  const item = Office.context.mailbox.item;
  showStatus("s-status", "Bezig…", "inf");
  try {
    await ewsMoveItem(item.itemId, singleFolder);
    showStatus("s-status", `✅ Verplaatst naar "${singleFolder.name}"`, "ok");
    setTimeout(() => showScreen("home"), 1500);
  } catch (e) {
    showStatus("s-status", "❌ " + e.message, "err");
  }
}

// ══════════════════════════════════════════════════════════════
//  BULK MODE
// ══════════════════════════════════════════════════════════════
async function startBulkMode() {
  showScreen("bulk");
  show("bulk-loading"); hide("bulk-card"); hide("bulk-done"); hide("bulk-error");
  bulkStats = { actie: 0, overgeslagen: 0, verplaatst: 0 };
  bulkIndex = 0;

  try {
    const [emails, folders] = await Promise.all([getInboxEmails(), getFolders()]);
    allFolders  = folders;
    bulkEmails  = emails;

    if (emails.length === 0) {
      showBulkDone("Geen e-mails gevonden in het Postvak In.");
      return;
    }
    hide("bulk-loading"); show("bulk-card");
    loadBulkEmail(0);
  } catch (e) {
    setEl("b-err-msg", "❌ " + e.message);
    hide("bulk-loading"); show("bulk-error");
  }
}

async function loadBulkEmail(idx) {
  bulkAiResult      = null;
  bulkSelectedFolder= null;
  hideFolderPanel();
  hide("b-status");

  const total = bulkEmails.length;

  if (idx >= total) {
    showBulkDone(`${total} mails doorlopen — ${bulkStats.actie} gemarkeerd, ${bulkStats.verplaatst} verplaatst, ${bulkStats.overgeslagen} overgeslagen.`);
    return;
  }

  const email = bulkEmails[idx];

  // Progress
  setEl("bulk-progress-label", `Mail ${idx + 1} van ${total}`);
  document.getElementById("bulk-progress-bar").style.width = Math.round(((idx) / total) * 100) + "%";

  // Show email header immediately
  setEl("b-from",    esc(email.from));
  setEl("b-subject", esc(email.subject));
  setBadge("b-badge", null);
  setEl("b-summary", '<span class="spinner"></span> AI analyseert…');
  disableBulkButtons(true);

  // Load body + call AI
  try {
    const body = await getBodyByItemId(email.id, email.changeKey);
    const ai   = await callClaude(email.subject, email.from, body, allFolders);
    bulkAiResult = ai;

    setBadge("b-badge", ai.action_type);
    setEl("b-summary", esc(ai.action_summary));

    // Pre-fill folder panel
    bulkSelectedFolder = allFolders.find(f => matchFolder(f, ai.suggested_folder)) || null;
    setEl("b-folder",   esc(ai.suggested_folder));
    setConfBar("b-conf-bar", "b-conf-lbl", ai.confidence);
    setEl("b-reason",   esc(ai.folder_reason));
    renderFolderList("b-list", allFolders, bulkSelectedFolder, (f) => { bulkSelectedFolder = f; });

    // Update action button label
    const actionLabel = actionTypeLabel(ai.action_type);
    document.getElementById("b-btn-action").innerHTML = `✅ ${actionLabel}`;

    disableBulkButtons(false);
  } catch (e) {
    setEl("b-summary", "❌ AI-fout: " + esc(e.message) + " — kies zelf een actie.");
    disableBulkButtons(false);
  }
}

function bulkDoAction() {
  // Flag/mark email for follow-up, then advance
  const email = bulkEmails[bulkIndex];
  ewsFlagItem(email.id, email.changeKey)
    .then(() => {
      bulkStats.actie++;
      advanceBulk(`✅ Gemarkeerd voor opvolging`);
    })
    .catch(e => {
      showStatus("b-status", "⚠️ Markeren mislukt: " + e.message, "err");
    });
}

function bulkSkip() {
  bulkStats.overgeslagen++;
  advanceBulk(null);
}

function bulkToggleFolder() {
  const panel = document.getElementById("folder-panel");
  panel.style.display = panel.style.display === "none" ? "block" : "none";
}

function hideFolderPanel() {
  document.getElementById("folder-panel").style.display = "none";
}

async function bulkConfirmMove() {
  if (!bulkSelectedFolder) { showStatus("b-status", "Selecteer een map.", "err"); return; }
  const email = bulkEmails[bulkIndex];
  try {
    await ewsMoveItem(email.id, bulkSelectedFolder);
    bulkStats.verplaatst++;
    hideFolderPanel();
    advanceBulk(`📁 Verplaatst naar "${bulkSelectedFolder.name}"`);
  } catch (e) {
    showStatus("b-status", "❌ " + e.message, "err");
  }
}

function advanceBulk(statusMsg) {
  if (statusMsg) showStatus("b-status", statusMsg, "ok");
  bulkIndex++;
  setTimeout(() => loadBulkEmail(bulkIndex), statusMsg ? 900 : 0);
}

function showBulkDone(statsText) {
  hide("bulk-card");
  setEl("bulk-done-stats", statsText);
  show("bulk-done");
}

function disableBulkButtons(on) {
  ["b-btn-action","b-btn-skip","b-btn-move"].forEach(id => {
    document.getElementById(id).disabled = on;
  });
}

// ══════════════════════════════════════════════════════════════
//  EWS HELPERS
// ══════════════════════════════════════════════════════════════

/* Get all mail folders */
function getFolders() {
  return new Promise((resolve, reject) => {
    const xml = ewsEnvelope(`
      <FindFolder Traversal="Deep"
        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
        <FolderShape>
          <t:BaseShape>IdOnly</t:BaseShape>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="folder:DisplayName"/>
            <t:FieldURI FieldURI="folder:FolderClass"/>
          </t:AdditionalProperties>
        </FolderShape>
        <ParentFolderIds>
          <t:DistinguishedFolderId Id="msgfolderroot"/>
        </ParentFolderIds>
      </FindFolder>`);

    Office.context.mailbox.makeEwsRequestAsync(xml, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded)
        return reject(new Error("Mappen ophalen mislukt: " + result.error.message));
      try {
        const doc     = parseXml(result.value);
        const NS_T    = "http://schemas.microsoft.com/exchange/services/2006/types";
        const items   = doc.getElementsByTagNameNS(NS_T, "Folder");
        const SKIP    = ["deleteditems","junkemail","drafts","outbox","sentitems",
                         "searchfolders","syncissues","conflicts","localfailures",
                         "serverfailures","recoverableitemsdeletions"];
        const folders = [];
        for (const f of items) {
          const nameEl  = f.getElementsByTagNameNS(NS_T, "DisplayName")[0];
          const idEl    = f.getElementsByTagNameNS(NS_T, "FolderId")[0];
          const clsEl   = f.getElementsByTagNameNS(NS_T, "FolderClass")[0];
          if (!nameEl || !idEl) continue;
          const name    = nameEl.textContent.trim();
          const cls     = clsEl ? clsEl.textContent : "IPF.Note";
          if (!cls.includes("IPF.Note")) continue;
          if (SKIP.some(s => name.toLowerCase().includes(s))) continue;
          if (!name || name === "Top of Information Store") continue;
          folders.push({ id: idEl.getAttribute("Id"), changeKey: idEl.getAttribute("ChangeKey"), name });
        }
        folders.sort((a, b) => a.name.localeCompare(b.name, "nl"));
        if (!folders.length) return reject(new Error("Geen mappen gevonden."));
        resolve(folders);
      } catch (e) { reject(e); }
    });
  });
}

/* Get inbox email headers */
function getInboxEmails(max = 60) {
  return new Promise((resolve, reject) => {
    const xml = ewsEnvelope(`
      <FindItem Traversal="Shallow"
        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
        <ItemShape>
          <t:BaseShape>IdOnly</t:BaseShape>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="item:Subject"/>
            <t:FieldURI FieldURI="message:From"/>
            <t:FieldURI FieldURI="item:DateTimeReceived"/>
          </t:AdditionalProperties>
        </ItemShape>
        <IndexedPageItemView MaxEntriesReturned="${max}" Offset="0" BasePoint="Beginning"/>
        <ParentFolderIds>
          <t:DistinguishedFolderId Id="inbox"/>
        </ParentFolderIds>
      </FindItem>`);

    Office.context.mailbox.makeEwsRequestAsync(xml, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded)
        return reject(new Error("Inbox ophalen mislukt: " + result.error.message));
      try {
        const NS_T  = "http://schemas.microsoft.com/exchange/services/2006/types";
        const doc   = parseXml(result.value);
        const msgs  = doc.getElementsByTagNameNS(NS_T, "Message");
        const items = [];
        for (const m of msgs) {
          const idEl      = m.getElementsByTagNameNS(NS_T, "ItemId")[0];
          const subjEl    = m.getElementsByTagNameNS(NS_T, "Subject")[0];
          const fromEl    = m.getElementsByTagNameNS(NS_T, "EmailAddress")[0];
          const fromName  = m.getElementsByTagNameNS(NS_T, "Name")[0];
          if (!idEl) continue;
          items.push({
            id:        idEl.getAttribute("Id"),
            changeKey: idEl.getAttribute("ChangeKey"),
            subject:   subjEl  ? subjEl.textContent  : "(geen onderwerp)",
            from:      fromName ? fromName.textContent : (fromEl ? fromEl.textContent : "onbekend"),
          });
        }
        resolve(items);
      } catch (e) { reject(e); }
    });
  });
}

/* Get body of specific item via EWS GetItem */
function getBodyByItemId(itemId, changeKey) {
  return new Promise((resolve, reject) => {
    const xml = ewsEnvelope(`
      <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
        <ItemShape>
          <t:BaseShape>IdOnly</t:BaseShape>
          <t:BodyType>Text</t:BodyType>
          <t:AdditionalProperties>
            <t:FieldURI FieldURI="item:Body"/>
          </t:AdditionalProperties>
        </ItemShape>
        <ItemIds>
          <t:ItemId Id="${itemId}" ChangeKey="${changeKey}"/>
        </ItemIds>
      </GetItem>`);

    Office.context.mailbox.makeEwsRequestAsync(xml, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded)
        return resolve("(body niet beschikbaar)");
      try {
        const NS_T = "http://schemas.microsoft.com/exchange/services/2006/types";
        const doc  = parseXml(result.value);
        const body = doc.getElementsByTagNameNS(NS_T, "Body")[0];
        resolve(body ? trunc(body.textContent, 1500) : "(leeg)");
      } catch { resolve("(body niet beschikbaar)"); }
    });
  });
}

/* Get body from currently open item (single mode) */
function getBody(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, {}, (r) => {
      resolve(r.status === Office.AsyncResultStatus.Succeeded
        ? trunc(r.value, 1500) : "(body niet beschikbaar)");
    });
  });
}

/* Move item to folder */
function ewsMoveItem(itemId, folder) {
  return new Promise((resolve, reject) => {
    const xml = ewsEnvelope(`
      <MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
        <ToFolderId>
          <t:FolderId Id="${folder.id}" ChangeKey="${folder.changeKey}"/>
        </ToFolderId>
        <ItemIds>
          <t:ItemId Id="${itemId}"/>
        </ItemIds>
      </MoveItem>`);

    Office.context.mailbox.makeEwsRequestAsync(xml, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded)
        return reject(new Error("Verplaatsen mislukt: " + result.error.message));
      const doc = parseXml(result.value);
      const rc  = doc.querySelector("MoveItemResponseMessage")?.getAttribute("ResponseClass");
      if (rc === "Success") resolve();
      else {
        const msg = doc.querySelector("MessageText")?.textContent || "Onbekende fout";
        reject(new Error(msg));
      }
    });
  });
}

/* Flag item for follow-up */
function ewsFlagItem(itemId, changeKey) {
  return new Promise((resolve, reject) => {
    const xml = ewsEnvelope(`
      <UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite"
        xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
        <ItemChanges>
          <t:ItemChange>
            <t:ItemId Id="${itemId}" ChangeKey="${changeKey}"/>
            <t:Updates>
              <t:SetItemField>
                <t:FieldURI FieldURI="item:Flag"/>
                <t:Message>
                  <t:Flag>
                    <t:FlagStatus>Flagged</t:FlagStatus>
                  </t:Flag>
                </t:Message>
              </t:SetItemField>
            </t:Updates>
          </t:ItemChange>
        </ItemChanges>
      </UpdateItem>`);

    Office.context.mailbox.makeEwsRequestAsync(xml, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded)
        return reject(new Error("Markeren mislukt: " + result.error.message));
      resolve();
    });
  });
}

function ewsEnvelope(body) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>
  <soap:Body>${body}</soap:Body>
</soap:Envelope>`;
}

// ══════════════════════════════════════════════════════════════
//  CLAUDE API
// ══════════════════════════════════════════════════════════════
async function callClaude(subject, from, body, folders) {
  const apiKey   = localStorage.getItem("claudeApiKey");
  const ctx      = localStorage.getItem("extraContext") || "";
  const folderList = folders.map(f => f.name).join("\n");

  const prompt = `Je bent een e-mailorganisatie-assistent voor een professionele gebruiker.
${ctx ? `Context: ${ctx}` : ""}

Beschikbare mappen:
${folderList}

Analyseer deze e-mail:
Van: ${from}
Onderwerp: ${subject}
Inhoud (fragment): ${body}

Geef je antwoord UITSLUITEND als JSON, zonder uitleg of markdown:
{
  "action_summary": "<max 1 zin: wat vraagt deze mail van de ontvanger?>",
  "action_type": "<reply|task|review|fyi|meeting>",
  "suggested_folder": "<exacte naam uit de mappenlijst>",
  "confidence": <0-100>,
  "folder_reason": "<max 1 zin waarom deze map>"
}`;

  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true"
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 400,
      messages: [{ role: "user", content: prompt }]
    })
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error("Claude API: " + (err.error?.message || resp.statusText));
  }

  const data = await resp.json();
  const raw  = data.content?.[0]?.text || "{}";
  try {
    return JSON.parse(raw.replace(/```json|```/g, "").trim());
  } catch {
    return { action_summary: "Kon niet analyseren.", action_type: "fyi",
             suggested_folder: folders[0]?.name || "Inbox", confidence: 30, folder_reason: "" };
  }
}

// ══════════════════════════════════════════════════════════════
//  UI HELPERS
// ══════════════════════════════════════════════════════════════
function renderFolderList(listId, folders, active, onSelect) {
  const list = document.getElementById(listId);
  list.innerHTML = "";
  for (const folder of folders) {
    const div = document.createElement("div");
    div.className = "folder-item" + (active && active.id === folder.id ? " active" : "");
    div.innerHTML = `<span>📁</span>${esc(folder.name)}`;
    div.onclick = () => {
      list.querySelectorAll(".folder-item").forEach(el => el.classList.remove("active"));
      div.classList.add("active");
      // update suggestion label if in bulk
      if (listId === "b-list") { setEl("b-folder", esc(folder.name)); }
      if (listId === "s-list") { setEl("s-folder", esc(folder.name)); }
      onSelect(folder);
    };
    list.appendChild(div);
  }
}

function filterFolders(listId, query) {
  const q       = query.toLowerCase();
  const onSelect = listId === "b-list"
    ? (f) => { bulkSelectedFolder = f; }
    : (f) => { singleFolder = f; };
  const active  = listId === "b-list" ? bulkSelectedFolder : singleFolder;
  const filtered = q ? allFolders.filter(f => f.name.toLowerCase().includes(q)) : allFolders;
  renderFolderList(listId, filtered, active, onSelect);
}

function matchFolder(folder, name) {
  if (!name) return false;
  const n = name.toLowerCase();
  return folder.name.toLowerCase() === n || folder.name.toLowerCase().includes(n);
}

const ACTION_LABELS = {
  reply:   "Beantwoorden",
  task:    "Taak aanmaken",
  review:  "Beoordelen",
  fyi:     "Ter kennisgeving",
  meeting: "Meeting actie"
};

function actionTypeLabel(type) { return ACTION_LABELS[type] || "Actie"; }

const BADGE_CLASSES = {
  reply: "badge-reply", task: "badge-task", review: "badge-review",
  fyi: "badge-fyi", meeting: "badge-meeting"
};

function setBadge(id, type) {
  const el = document.getElementById(id);
  el.className = "badge " + (BADGE_CLASSES[type] || "badge-fyi");
  el.textContent = actionTypeLabel(type);
}

function setConfBar(barId, lblId, conf) {
  const c = Math.max(0, Math.min(100, conf || 0));
  document.getElementById(barId).style.width = c + "%";
  setEl(lblId, `Zekerheid: ${c}%`);
}

function showStatus(id, msg, type) {
  const el = document.getElementById(id);
  el.className = `status status-${type}`;
  el.innerHTML = msg;
  el.style.display = "block";
}

function setEl(id, html)  { const e = document.getElementById(id); if (e) e.innerHTML = html; }
function show(id)         { const e = document.getElementById(id); if (e) e.style.display = "block"; }
function hide(id)         { const e = document.getElementById(id); if (e) e.style.display = "none"; }
function esc(s)           { return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
function trunc(s, n)      { return s && s.length > n ? s.slice(0, n) + "…" : s; }
function parseXml(str)    { return new DOMParser().parseFromString(str, "text/xml"); }
