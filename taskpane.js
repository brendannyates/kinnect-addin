/* Kinnect: front desk <-> resource notifications in a shared workbook (Excel on the web), carried over from the Floot kin-nect app.
   A resource is a clinician or a device (MagVenture, BrainsWay). Messages are rows in the KinnectLog table ("Kinnect Log" sheet);
   resources, with their Zoom and Nudge AI links and default room, are rows in KinnectResources ("Resources" sheet).
   Today's NextGen report is pasted on "Appt List". One-click alerts carry no patient identifier; a LASFIR code is optional.
   Nothing leaves the workbook; this page holds no PHI. */
"use strict";
const DEMO = /[?&]demo\b/.test(location.search);
const LOG_SHEET = "Kinnect Log", LOG_TABLE = "KinnectLog", RES_SHEET = "Resources", RES_TABLE = "KinnectResources", APPT_SHEET = "Appt List";
const LOG_COLS = ["ID", "Sent", "From", "To", "Patient", "Status", "Room", "Note", "Acked by", "Acked at"];
const RES_COLS = ["Resource", "Zoom URL", "Nudge URL", "Room"];
const L = Object.fromEntries(LOG_COLS.map((c, i) => [c, i]));
const FD = "Front desk", ALL = "All resources";
const ARRIVED = "Patient arrived", TELE = "Telehealth patient arrived";
const FD_STATUS = [ARRIVED, TELE, "Roomed", "Running late", "No-show", "Cancelled", "Checked out", "Message"];
const RES_STATUS = ["Ready for next patient", "Send patient back", "Ready for checkout", "Need help", "Running behind", "Message"];
const SAYS = { [ARRIVED]: "Your patient has arrived.", [TELE]: "Your telehealth patient has arrived." };
const POLL_MS = 10000, DEDUP_MS = 5000;
const LASFIR = /^[A-Z]{2,6}$/;   // short names give shorter codes (Ng, Al = NGAL)
/* the four kin-nect sound slots: drop bing/boom/bop/bow (.mp3, .wav, .ogg or .m4a) into assets/sounds; a missing file falls back to the built-in chime */
const SOUNDS = ["bing", "boom", "bop", "bow"], SOUND_EXTS = ["mp3", "wav", "ogg", "m4a"];
const DEV_DEFAULT = { receive: true, sound: "bing", volume: 0.8, desktop: false, ptCode: false };

const state = { ready: false, log: [], res: [], appts: [], apptNote: "", view: "main", me: load("kinnect.me"), dev: { ...DEV_DEFAULT, ...(load("kinnect.dev") || {}) }, alerted: new Set(), first: true, sent: new Map() };
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
function load(k) { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } }
function save(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} }
const setDev = patch => { state.dev = { ...state.dev, ...patch }; save("kinnect.dev", state.dev); };
const pad = n => String(n).padStart(2, "0");
const stamp = (d = new Date()) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
const today = () => stamp().slice(0, 10);
const hm = s => { const t = String(s).slice(11, 16); if (!t) return ""; const [h, m] = t.split(":").map(Number); return `${(h + 11) % 12 + 1}:${pad(m)} ${h < 12 ? "AM" : "PM"}`; };
const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const https = u => /^https:\/\/\S+$/i.test(String(u || "").trim());
const myName = () => state.me && (state.me.role === "fd" ? FD : state.me.name);

let toastTimer;
function toast(msg, err) { const t = $("toast"); t.textContent = msg; t.className = err ? "err" : ""; t.style.display = "block"; clearTimeout(toastTimer); toastTimer = setTimeout(() => t.style.display = "none", err ? 8000 : 3500); }
const fail = e => { console.error(e); toast(e.message || String(e), true); };

/* "Last, First" or "First Last" to LASFIR; a code passes through */
function toLasfir(v) {
  const s = String(v || "").trim().toUpperCase(); if (LASFIR.test(s)) return s;
  const clean = x => x.replace(/[^A-Z]/g, "");
  let last, first;
  if (s.includes(",")) [last, first] = s.split(",").map(x => clean(x.trim().split(/\s+/)[0] || ""));
  else { const w = s.split(/\s+/).map(clean).filter(Boolean); first = w[0]; last = w[w.length - 1]; }
  const code = (last || "").slice(0, 3) + (first || "").slice(0, 3);
  return LASFIR.test(code) ? code : "";
}

/* ---------- workbook store (Excel) and an in-memory stand-in for ?demo ---------- */
const excelStore = {
  async read() {
    let out = null;
    await Excel.run(async ctx => {
      const lt = ctx.workbook.tables.getItemOrNullObject(LOG_TABLE), rt = ctx.workbook.tables.getItemOrNullObject(RES_TABLE), aw = ctx.workbook.worksheets.getItemOrNullObject(APPT_SHEET);
      await ctx.sync();
      if (lt.isNullObject || rt.isNullObject) { out = { ready: false }; return; }
      const lb = lt.getDataBodyRange(), rb = rt.getDataBodyRange(), au = aw.isNullObject ? null : aw.getUsedRangeOrNullObject(true);
      lb.load("values"); rb.load("values"); if (au) au.load("values");
      await ctx.sync();
      out = { ready: true, log: lb.values, res: rb.values, appt: au && !au.isNullObject ? au.values : null };
    });
    return out;
  },
  async setup() {
    await Excel.run(async ctx => {
      const want = [[LOG_SHEET, LOG_TABLE, LOG_COLS], [RES_SHEET, RES_TABLE, RES_COLS]];
      const tbl = want.map(([, t]) => ctx.workbook.tables.getItemOrNullObject(t)), ws = want.map(([s]) => ctx.workbook.worksheets.getItemOrNullObject(s));
      await ctx.sync();
      want.forEach(([s, t, cols], i) => {
        if (!tbl[i].isNullObject) return;
        const sh = ws[i].isNullObject ? ctx.workbook.worksheets.add(s) : ws[i];
        const hdr = sh.getRange("A1").getResizedRange(0, cols.length - 1); hdr.values = [cols];
        const nt = sh.tables.add(hdr, true); nt.name = t;
        if (t === LOG_TABLE) nt.getDataBodyRange().numberFormat = [cols.map(() => "@")];   // keep IDs and times as typed text
        hdr.format.autofitColumns();
      });
      await ctx.sync();
    });
  },
  async append(row) {
    await Excel.run(async ctx => { ctx.workbook.tables.getItem(LOG_TABLE).rows.add(null, [row]); await ctx.sync(); });
  },
  async ack(id, by) {
    await Excel.run(async ctx => {
      const b = ctx.workbook.tables.getItem(LOG_TABLE).getDataBodyRange(); b.load("values"); await ctx.sync();
      const i = b.values.findIndex(r => r[L.ID] === id); if (i < 0) throw new Error("That message is no longer in the log.");
      b.getCell(i, L["Acked by"]).getResizedRange(0, 1).values = [[by, stamp()]]; await ctx.sync();
    });
  },
  async saveLinks(name, zoom, nudge) {
    await Excel.run(async ctx => {
      const t = ctx.workbook.tables.getItem(RES_TABLE), b = t.getDataBodyRange(); b.load("values"); await ctx.sync();
      const i = b.values.findIndex(r => String(r[0]).trim() === name);
      if (i >= 0) b.getCell(i, 1).getResizedRange(0, 1).values = [[zoom, nudge]];
      else { const blank = b.values.findIndex(r => r.every(v => v === "")); if (blank >= 0) b.getRow(blank).values = [[name, zoom, nudge, ""]]; else t.rows.add(null, [[name, zoom, nudge, ""]]); }
      await ctx.sync();
    });
  },
  watch(cb) {
    Excel.run(async ctx => { ctx.workbook.tables.getItem(LOG_TABLE).onChanged.add(async () => cb()); await ctx.sync(); }).catch(e => console.warn("change events unavailable", e));
  }
};
const mem = {   // fictional
  log: [["d1", `${today()} 08:52:10`, FD, "Dr. Rivera", "", ARRIVED, "", "", "", ""], ["d2", `${today()} 08:55:41`, "Dr. Rivera", FD, "", "Ready for next patient", "", "", FD, `${today()} 08:56:02`]],
  res: [["Dr. Rivera", "https://zoom.us/j/0000000000", "https://example.com/nudge", "2"], ["Dr. Okafor", "", "", "3"], ["MagVenture", "", "", "TMS 1"]],
  appt: [["Appt Time", "Patient", "Provider", "Status"], [0.375, "Doe, Jane", "Dr. Rivera", "Scheduled"], [0.4167, "Roe, Rich", "Dr. Okafor", "Scheduled"]]
};
const demoStore = {
  async read() { return { ready: true, log: mem.log.map(r => [...r]), res: mem.res.map(r => [...r]), appt: mem.appt }; },
  async setup() {}, async append(row) { mem.log.push(row); },
  async ack(id, by) { const r = mem.log.find(r => r[L.ID] === id); r[L["Acked by"]] = by; r[L["Acked at"]] = stamp(); },
  async saveLinks(name, zoom, nudge) { const r = mem.res.find(r => r[0] === name); if (r) { r[1] = zoom; r[2] = nudge; } else mem.res.push([name, zoom, nudge, ""]); },
  watch() {}
};
const store = DEMO ? demoStore : excelStore;

/* ---------- read and parse ---------- */
/* NextGen report on Appt List: finds the header row by name, so column order doesn't matter. Adjust APPT_HDR once the report exists. */
const APPT_HDR = { time: /appt.*time|^time|start/i, patient: /patient|lasfir|^name/i, res: /provider|resource|clinician|rendering/i, status: /status/i, date: /appt.*date|^date/i };
function parseAppts(v) {
  if (!v || !v.length) return { rows: [], note: `Paste today's NextGen report on the "${APPT_SHEET}" sheet to list arrivals here.` };
  const h = v.findIndex(r => r.some(c => APPT_HDR.patient.test(String(c))) && r.some(c => APPT_HDR.res.test(String(c))));
  if (h < 0) return { rows: [], note: `Couldn't find Patient and Provider headers on "${APPT_SHEET}".` };
  const col = Object.fromEntries(Object.entries(APPT_HDR).map(([k, re]) => [k, v[h].findIndex(c => re.test(String(c)))]));
  const tSerial = Math.floor((new Date() - new Date(1899, 11, 30)) / 864e5);
  const rows = v.slice(h + 1).map(r => {
    const t = col.time >= 0 ? r[col.time] : "", frac = typeof t === "number" ? t % 1 : null;
    return { lasfir: toLasfir(r[col.patient]), res: String(r[col.res] ?? "").trim(), status: col.status >= 0 ? String(r[col.status]) : "",
      date: col.date >= 0 && typeof r[col.date] === "number" ? Math.floor(r[col.date]) : null,
      min: frac != null ? Math.round(frac * 1440) : null, time: frac != null ? `${(Math.floor(frac * 24) + 11) % 12 + 1}:${pad(Math.round(frac * 1440) % 60)} ${frac < .5 ? "AM" : "PM"}` : String(t) };
  }).filter(r => r.lasfir && r.res && (r.date == null || r.date === tSerial) && !/cancel/i.test(r.status));
  rows.sort((a, b) => (a.min ?? 9999) - (b.min ?? 9999));
  return { rows, note: "" };
}
/* Excel can turn a typed timestamp into a date serial: read either back as local "yyyy-mm-dd hh:mm:ss" */
const when = v => { if (typeof v !== "number") return String(v); const d = new Date(Math.round((v - 25569) * 864e5)); return stamp(new Date(d.getTime() + d.getTimezoneOffset() * 6e4)); };
const toMsg = r => ({ id: String(r[L.ID]), sent: when(r[L.Sent]), from: String(r[L.From]), to: String(r[L.To]), pt: String(r[L.Patient]), status: String(r[L.Status]), room: String(r[L.Room]), note: String(r[L.Note]), ackBy: String(r[L["Acked by"]]), ackAt: when(r[L["Acked at"]]) });
async function refresh() {
  try {
    const d = await store.read(); state.ready = d.ready;
    if (d.ready) {
      state.log = d.log.filter(r => r[L.ID] !== "").map(toMsg).filter(m => m.sent.startsWith(today()));
      state.res = d.res.filter(r => String(r[0]).trim()).map(r => ({ name: String(r[0]).trim(), zoom: String(r[1]).trim(), nudge: String(r[2]).trim(), room: String(r[3]).trim() }));
      const a = parseAppts(d.appt); state.appts = a.rows; state.apptNote = a.note;
    }
    render(); alertNew();
  } catch (e) { fail(e); }
}

/* ---------- sound: the chosen kin-nect sound file, else a built-in chime ---------- */
const soundUrl = {};
function findSound(id) {
  if (id in soundUrl) return Promise.resolve(soundUrl[id]);
  return new Promise(done => {
    const a = new Audio(); let i = 0, over = false;
    const finish = (u, keep = true) => { if (over) return; over = true; clearTimeout(t); a.onloadedmetadata = a.onerror = null; if (keep) soundUrl[id] = u; done(u); };
    const next = () => { if (i >= SOUND_EXTS.length) return finish(null); a.src = `assets/sounds/${id}.${SOUND_EXTS[i++]}`; a.load(); };
    const t = setTimeout(() => finish(null, false), 3000);   // slow network: chime this time, look again next time
    a.preload = "metadata"; a.onloadedmetadata = () => finish(a.src); a.onerror = next; next();
  });
}
let audioCtx;
function chime(vol) {
  try {
    audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
    const c = audioCtx, o = c.createOscillator(), g = c.createGain(); o.frequency.value = 880;
    g.gain.setValueAtTime(.25 * vol, c.currentTime); g.gain.exponentialRampToValueAtTime(.001, c.currentTime + .4); o.connect(g).connect(c.destination); o.start(); o.stop(c.currentTime + .4);
  } catch {}
}
const player = new Audio();
async function playSound(id = state.dev.sound) {
  const vol = Math.min(1, Math.max(0, Number(state.dev.volume) || 0));
  if (id === "none" || !vol) return;
  const u = SOUNDS.includes(id) ? await findSound(id) : null;
  if (!u) return chime(vol);
  player.src = u; player.volume = vol; player.currentTime = 0;
  player.play().catch(() => chime(vol));   // autoplay can be blocked until the pane has been clicked once
}
/* browsers only allow sound after a click in the page: the first click anywhere unlocks it */
document.addEventListener("click", () => { try { audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)(); audioCtx.resume(); } catch {} }, { once: true });

/* ---------- alerts for unacknowledged messages to me: sound, red banner, title count, optional desktop pop-up ---------- */
const mine = m => state.me && (state.me.role === "fd" ? m.to === FD : m.to === state.me.name || m.to === ALL);
const inbox = () => state.log.filter(mine).sort((a, b) => (!!a.ackBy - !!b.ackBy) || b.sent.localeCompare(a.sent));
function desktopPop(m) {
  if (!state.dev.desktop || !("Notification" in window) || Notification.permission !== "granted") return;
  try { new Notification("Kinnect", { body: `${SAYS[m.status] || m.status} (${m.from})`, tag: m.id }); } catch {}   // no patient code on a pop-up anyone walking by can read
}
function alertNew() {
  const open = state.dev.receive ? inbox().filter(m => !m.ackBy) : [], fresh = open.filter(m => !state.alerted.has(m.id));
  fresh.forEach(m => state.alerted.add(m.id));
  if (fresh.length && !state.first) { playSound(); fresh.forEach(desktopPop); }
  state.first = false;
  document.title = open.length ? `(${open.length}) Kinnect` : "Kinnect";
  $("banner").innerHTML = open.length ? `<div class="banner red">${open.length} new message${open.length === 1 ? "" : "s"}</div>` : !state.dev.receive && state.me ? `<div class="banner">Notifications are off on this computer. Turn them on in Settings.</div>` : "";
}

/* ---------- views ---------- */
/* re-rendering every poll must not lose what someone is typing or which sections they opened */
function render() {
  const keep = [...document.querySelectorAll("#main input, #main select")].map(x => [x.id, x.type === "checkbox" ? null : x.value]), open = [...document.querySelectorAll("#main details")].map(d => d.open);
  const focus = document.activeElement && document.activeElement.id, sel = focus && document.activeElement.selectionStart;
  paint();
  keep.forEach(([id, v]) => { const x = id && v != null && $(id); if (x && (x.tagName !== "SELECT" || [...x.options].some(o => o.value === v))) x.value = v; });
  document.querySelectorAll("#main details").forEach((d, i) => { if (i < open.length) d.open = open[i]; });
  if (focus && $(focus)) { $(focus).focus(); try { if (sel != null) $(focus).setSelectionRange(sel, sel); } catch {} }
}
function paint() {
  const me = state.me;
  $("who").innerHTML = me ? `<span class="who" data-act="who" title="Change who uses this computer">${esc(myName())}</span> <button class="sm" data-act="${state.view === "settings" ? "home" : "settings"}">${state.view === "settings" ? "Done" : "Settings"}</button>` : "";
  if (!state.ready) return $("main").innerHTML = `<section><p>This workbook doesn't have the Kinnect sheets yet.</p><button class="pri" data-act="setup">Set up workbook</button><p class="hint">Adds a "${LOG_SHEET}" sheet for messages and a "${RES_SHEET}" sheet for each clinician or device, with Zoom and Nudge AI links.</p></section>`;
  if (!me) return $("main").innerHTML = viewWho();
  $("main").innerHTML = state.view === "settings" ? viewSettings() : me.role === "fd" ? viewFD() : viewRes(me.name);
}
function viewWho() {
  const opts = state.res.map(c => `<option>${esc(c.name)}</option>`).join("");
  return `<section><h2>Who's using this computer?</h2><p class="sub">Saved on this computer only.</p>
    <button class="pri big" data-act="beFD" style="width:100%;margin-bottom:12px">Front desk</button>
    <label for="whoRes">Clinician or device</label><select id="whoRes"><option value="">Choose...</option>${opts}</select>
    <label for="whoNew">Or add one</label><input id="whoNew" placeholder="Name as it appears in NextGen">
    <div style="margin-top:8px"><button data-act="beRes">Use this computer for it</button></div></section>`;
}
function viewSettings() {
  const d = state.dev, perm = "Notification" in window ? Notification.permission : "unsupported";
  const opt = (v, l) => `<option value="${v}"${d.sound === v ? " selected" : ""}>${l}</option>`;
  const files = SOUNDS.map(id => `${id}: ${id in soundUrl ? (soundUrl[id] ? "found" : "missing, uses the chime") : "not checked yet"}`).join(" · ");
  return `<section>
    <div class="group"><h2>This computer</h2>
      <label class="chk"><input type="checkbox" id="sRecv"${d.receive ? " checked" : ""}> Notify me on this computer</label>
      <label for="sSound">Sound</label><select id="sSound">${SOUNDS.map(s => opt(s, s[0].toUpperCase() + s.slice(1))).join("")}${opt("chime", "Built-in chime")}${opt("none", "No sound")}</select>
      <label for="sVol">Volume</label><input type="range" id="sVol" min="0" max="1" step="0.05" value="${d.volume}">
      <div style="margin-top:6px"><button class="sm" data-act="testSound">Test sound</button></div>
      <div class="hint">${esc(files)}</div>
      <label class="chk"><input type="checkbox" id="sDesk"${d.desktop ? " checked" : ""}> Desktop pop-ups</label>
      <div class="hint">${perm === "granted" ? "Allowed in this browser." : perm === "denied" ? "Blocked in this browser's site settings." : perm === "unsupported" ? "Not available here." : "Excel may block the permission prompt inside the pane; the banner and sound still work."} Pop-ups never show the patient code.</div>
      ${state.me.role === "fd" ? `<label class="chk"><input type="checkbox" id="sPt"${d.ptCode ? " checked" : ""}> Add the patient code to one-click alerts from the appointment list</label><div class="hint">Off by default, matching kin-nect: "Patient arrived" says nothing about who. Turn on only if clinicians need to tell patients apart.</div>` : ""}
    </div></section>`;
}
function msgItem(m, canAck) {
  const who = mine(m) ? `from ${esc(m.from)}` : `to ${esc(m.to)}`;
  const ack = m.ackBy ? `<span class="chip">Seen by ${esc(m.ackBy)} ${hm(m.ackAt)}</span>` : canAck ? `<button class="sm pri" data-act="ack" data-id="${esc(m.id)}">Got it</button>` : `<span class="chip">Not seen yet</span>`;
  const head = mine(m) && SAYS[m.status] ? SAYS[m.status] : m.status;
  return `<div class="item${!m.ackBy && canAck ? " new" : m.ackBy ? " ack" : ""}"><b>${esc(head)}</b>${m.pt ? ` <span class="chip">${esc(m.pt)}</span>` : ""}${m.room ? ` <span class="chip">Room ${esc(m.room)}</span>` : ""}
    <div class="sub">${hm(m.sent)} ${who}${m.note ? ` · ${esc(m.note)}` : ""}</div>${ack}</div>`;
}
function feeds() {
  const inb = inbox(), sent = state.log.filter(m => !mine(m) && m.from === myName()).sort((a, b) => b.sent.localeCompare(a.sent));
  return `<div class="group"><h2>Inbox<span class="n">${inb.length}</span></h2>${inb.map(m => msgItem(m, !m.ackBy)).join("") || `<div class="none">Nothing yet today</div>`}</div>
    <div class="group"><h2>Sent today<span class="n">${sent.length}</span></h2>${sent.map(m => msgItem(m, false)).join("") || `<div class="none">Nothing sent today</div>`}</div>`;
}
const ptHint = `<div class="hint">Patient is optional. If needed, LASFIR only (last three + first three letters, e.g. SMIJAN). Don't type names or other details in the note.</div>`;
function viewFD() {
  const cards = state.res.map(r => `<div class="card"><b>${esc(r.name)}</b><div class="row" style="margin-top:6px">
      <button class="pri" data-act="quick" data-to="${esc(r.name)}" data-st="${ARRIVED}">Patient arrived</button>
      <button data-act="quick" data-to="${esc(r.name)}" data-st="${TELE}">Telehealth patient arrived</button></div></div>`).join("");
  const last = new Map(); state.log.filter(m => m.from === FD && m.pt).forEach(m => last.set(`${m.pt}|${m.to}`, m.status));
  const pt = a => state.dev.ptCode ? a.lasfir : "";
  const appts = state.appts.map(a => {
    const st = last.get(`${a.lasfir}|${a.res}`), room = (state.res.find(c => c.name === a.res) || {}).room || "";
    return `<div class="item"><b>${esc(a.time)}</b> <span class="chip">${esc(a.lasfir)}</span> <span class="sub">${esc(a.res)}</span>${st ? ` <span class="chip">${esc(st)}</span>` : ""}
      <div>${[[ARRIVED, "Arrived"], [TELE, "Telehealth"], ["Roomed", "Roomed"], ["Running late", "Running late"], ["No-show", "No-show"]].map(([s, l]) => `<button class="sm${s === ARRIVED && !st ? " pri" : ""}" data-act="quick" data-pt="${esc(pt(a))}" data-to="${esc(a.res)}" data-st="${esc(s)}" data-room="${esc(s === "Roomed" ? room : "")}">${l}</button>`).join("")}</div></div>`;
  }).join("");
  return `<section>
    <div class="group"><h2>Resources<span class="n">${state.res.length}</span></h2><p class="sub">One click notifies the resource. No patient details are sent.</p>${cards || `<div class="none">Add clinicians and devices on the "${RES_SHEET}" sheet.</div>`}</div>
    <div class="group"><h2>Today's appointments<span class="n">${state.appts.length}</span></h2>${appts || `<div class="none">${esc(state.apptNote || "No appointments left today")}</div>`}</div>
    <details class="group"><summary><b>Send a message</b></summary>${form([ALL, ...state.res.map(c => c.name)], FD_STATUS)}</details>
    ${feeds()}</section>`;
}
function viewRes(name) {
  const c = state.res.find(x => x.name === name) || { zoom: "", nudge: "" };
  const go = (label, u) => https(u) ? `<button class="pri big" data-act="open" data-url="${esc(u)}">${label}</button>` : `<button class="big" disabled title="Add the link under My links">${label}</button>`;
  return `<section>
    <div class="group"><div class="launch">${go("Join Zoom", c.zoom)}${go("Open Nudge AI", c.nudge)}</div>
      <details${https(c.zoom) || https(c.nudge) ? "" : " open"}><summary class="sub">My links</summary>
        <label for="lnZoom">Zoom room link</label><input id="lnZoom" value="${esc(c.zoom)}" placeholder="https://zoom.us/j/...">
        <label for="lnNudge">Nudge AI link</label><input id="lnNudge" value="${esc(c.nudge)}" placeholder="https://...">
        <div style="margin-top:6px"><button class="sm" data-act="links">Save links</button></div></details></div>
    <div class="group"><h2>Tell the front desk</h2>
      <div>${RES_STATUS.slice(0, 3).map(s => `<button class="sm" data-act="quick" data-to="${FD}" data-st="${esc(s)}">${s}</button>`).join("")}</div>
      <details style="margin-top:6px"><summary class="sub">More</summary>${form([FD], RES_STATUS)}</details></div>
    ${feeds()}</section>`;
}
function form(tos, statuses) {
  return `<div class="row"><div><label for="fTo">To</label><select id="fTo">${tos.map(t => `<option>${esc(t)}</option>`).join("")}</select></div>
      <div><label for="fSt">Status</label><select id="fSt">${statuses.map(s => `<option>${esc(s)}</option>`).join("")}</select></div></div>
    <div class="row"><div><label for="fPt">Patient (optional)</label><input id="fPt" maxlength="6" placeholder="SMIJAN" style="text-transform:uppercase"></div>
      <div><label for="fRoom">Room</label><input id="fRoom" maxlength="10"></div></div>
    <label for="fNote">Note</label><input id="fNote" maxlength="140">${ptHint}
    <div style="margin-top:6px"><button class="pri" data-act="send">Send</button></div>`;
}

/* ---------- actions ---------- */
async function send(to, status, pt, room, note) {
  pt = String(pt || "").trim().toUpperCase();
  if (pt && !LASFIR.test(pt)) return toast("Patient must be a LASFIR code (letters only), e.g. SMIJAN.", true);
  const from = myName(), key = `${from}|${to}|${status}|${pt}`, now = Date.now();
  // a double click, or two people pressing the same button, within a few seconds sends one alert (kin-nect's debounce)
  if (now - (state.sent.get(key) || 0) < DEDUP_MS || state.log.some(m => `${m.from}|${m.to}|${m.status}|${m.pt}` === key && now - new Date(m.sent.replace(" ", "T")).getTime() < DEDUP_MS)) return toast("Already notified moments ago.");
  state.sent.set(key, now);
  try { await store.append([newId(), stamp(), from, to, pt, status, String(room || "").trim(), String(note || "").trim(), "", ""]); ["fPt", "fRoom", "fNote"].forEach(id => { if ($(id)) $(id).value = ""; }); toast(`Notified ${to}.`); await refresh(); } catch (e) { state.sent.delete(key); fail(e); }
}
function openUrl(u) {
  if (!https(u)) return toast("That link must start with https://", true);
  const ok = window.open(u, "_blank", "noopener"); if (!ok && window.Office?.context?.ui?.openBrowserWindow) Office.context.ui.openBrowserWindow(u);
}
const ACT = {
  setup: async () => { try { await store.setup(); toast(`Workbook ready. Add clinicians and devices on the ${RES_SHEET} sheet, or let them add themselves.`); await refresh(); store.watch(soon); } catch (e) { fail(e); } },
  who: () => { state.me = null; state.view = "main"; save("kinnect.me", null); render(); },
  settings: () => { state.view = "settings"; render(); Promise.all(SOUNDS.map(findSound)).then(() => state.view === "settings" && render()); },
  home: () => { state.view = "main"; render(); },
  beFD: () => { state.me = { role: "fd" }; save("kinnect.me", state.me); state.first = true; refresh(); },
  beRes: async () => {
    const name = ($("whoRes").value || $("whoNew").value).trim(); if (!name) return toast("Choose or type a name.", true);
    if (!state.res.some(c => c.name === name)) { try { await store.saveLinks(name, "", ""); } catch (e) { return fail(e); } }
    state.me = { role: "res", name }; save("kinnect.me", state.me); state.first = true; refresh();
  },
  quick: el => send(el.dataset.to, el.dataset.st, el.dataset.pt, el.dataset.room, ""),
  send: () => send($("fTo").value, $("fSt").value, $("fPt").value, $("fRoom").value, $("fNote").value),
  ack: async el => { try { await store.ack(el.dataset.id, myName()); await refresh(); } catch (e) { fail(e); } },
  open: el => openUrl(el.dataset.url),
  testSound: () => playSound(),
  links: async () => {
    const z = $("lnZoom").value.trim(), n = $("lnNudge").value.trim();
    if ((z && !https(z)) || (n && !https(n))) return toast("Links must start with https://", true);
    try { await store.saveLinks(state.me.name, z, n); toast("Links saved."); await refresh(); } catch (e) { fail(e); }
  }
};
document.addEventListener("click", e => { const el = e.target.closest("[data-act]"); if (el && ACT[el.dataset.act]) ACT[el.dataset.act](el); });
document.addEventListener("change", async e => {
  const id = e.target.id;
  if (id === "sRecv") { setDev({ receive: e.target.checked }); alertNew(); }
  else if (id === "sSound") { setDev({ sound: e.target.value }); playSound(); }
  else if (id === "sVol") { setDev({ volume: Number(e.target.value) }); playSound(); }
  else if (id === "sPt") setDev({ ptCode: e.target.checked });
  else if (id === "sDesk") {
    let on = e.target.checked;
    if (on && "Notification" in window && Notification.permission === "default") { try { on = (await Notification.requestPermission()) === "granted"; } catch { on = false; } }
    if (on && !("Notification" in window && Notification.permission === "granted")) { on = false; toast("This browser didn't allow desktop pop-ups here. The banner and sound still work.", true); }
    setDev({ desktop: on }); render();
  }
});

/* ---------- start: change events from the log table plus a light poll as a fallback ---------- */
let soonT; const soon = () => { clearTimeout(soonT); soonT = setTimeout(refresh, 400); };   // a burst of table changes = one read
async function start() {
  if (state.me && state.me.role === "clin") state.me.role = "res";   // saved by version 0.1
  await refresh();
  if (state.ready) store.watch(soon);
  setInterval(() => { if (!document.hidden) refresh(); }, POLL_MS);
}
if (DEMO || typeof Office === "undefined") start(); else Office.onReady(start);
