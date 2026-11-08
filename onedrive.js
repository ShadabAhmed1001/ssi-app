// SSI Production, standalone build.
// Replaces the Claude page runtime (window.claude) with:
//  - a store kept on this phone (IndexedDB), so the app works without internet
//  - OneDrive sync: each phone entry is its own small file in "<entries folder>/app/entries",
//    all pending entries are also written as phone-entries-<phone>.csv for Excel's Import from phone,
//    and the Excel workbook + database are read from OneDrive to refresh the app's data.
(function () {
  window.SSI_STANDALONE = true;
  const CFG0 = window.SSI_CONFIG || {};
  const LSK = "ssi-od-settings";
  const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
  const cfg = () => Object.assign({}, CFG0, lsGet(LSK, {}));
  const SYNCED = ["entries", "audits", "people"];

  // ---------- local store ----------
  const DBN = "ssi-standalone", ST = "docs", META = "meta";
  const idb = new Promise((ok, bad) => {
    const r = indexedDB.open(DBN, 1);
    r.onupgradeneeded = () => { r.result.createObjectStore(ST); r.result.createObjectStore(META); };
    r.onsuccess = () => ok(r.result); r.onerror = () => bad(r.error);
  });
  const C = {}, subs = {};
  let meta = {};
  const ready = idb.then((d) => Promise.all([
    new Promise((ok) => { const t = d.transaction(ST).objectStore(ST).openCursor(); t.onsuccess = () => { const cur = t.result; if (!cur) return ok(); const [c, id] = String(cur.key).split("\u0001"); (C[c] = C[c] || {})[id] = cur.value; cur.continue(); }; t.onerror = () => ok(); }),
    new Promise((ok) => { const t = d.transaction(META).objectStore(META).openCursor(); t.onsuccess = () => { const cur = t.result; if (!cur) return ok(); meta[cur.key] = cur.value; cur.continue(); }; t.onerror = () => ok(); }),
  ]));
  const put = (store, k, v) => idb.then((d) => new Promise((ok, bad) => { const tx = d.transaction(store, "readwrite"), s = tx.objectStore(store); v === undefined ? s.delete(k) : s.put(v, k); tx.oncomplete = ok; tx.onerror = () => bad(tx.error); }));
  const setMeta = (k, v) => { if (v === undefined) delete meta[k]; else meta[k] = v; return put(META, k, v); };
  const col = (c) => (C[c] = C[c] || {});
  const snap = (c) => ({ docs: Object.keys(col(c)).map((id) => ({ id, data: () => col(c)[id] })), metadata: { fromCache: false } });
  const fire = (c) => setTimeout(() => (subs[c] || []).forEach((f) => { try { f(snap(c)); } catch (e) { console.error(e); } }), 0);
  let n = 0;
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8) + (n++).toString(36);
  const dirty = () => meta.dirty || {};
  async function write(c, id, v, fromRemote) {
    if (v === undefined) delete col(c)[id]; else col(c)[id] = v;
    await put(ST, c + "\u0001" + id, v);
    if (!fromRemote && SYNCED.includes(c)) { const d = Object.assign({}, dirty()); d[c + "/" + id] = 1; await setMeta("dirty", d); kick(); }
    fire(c);
  }
  const stamp = (o) => Object.assign(o, { _at: new Date().toISOString(), _by: me.id });
  const ref = (c, id) => ({ id,
    set: async (b) => write(c, id, stamp(JSON.parse(JSON.stringify(b)))),
    update: async (b) => write(c, id, stamp(Object.assign({}, col(c)[id] || {}, JSON.parse(JSON.stringify(b))))),
    delete: async () => write(c, id, undefined),
    get: async () => ({ exists: id in col(c), data: () => col(c)[id] }) });
  const dbApi = {
    collection: (c) => ({ doc: (id) => ref(c, id || newId()), onSnapshot: (f) => { (subs[c] = subs[c] || []).push(f); ready.then(() => f(snap(c))); return () => {}; } }),
    doc: (p) => { const i = p.lastIndexOf("/"); return ref(p.slice(0, i), p.slice(i + 1)); },
  };

  // ---------- who is using it ----------
  const me = { id: lsGet("ssi-me", { id: "local" }).id, name: lsGet("ssi-me", { name: "" }).name || cfg().userName || "" };
  const userApi = {
    id: async () => me.id, me: async () => ({ id: me.id, name: me.name }),
    can: async () => true, canEdit: async () => true, isOwner: () => true,
    profiles: async (ids) => Object.fromEntries([].concat(ids).map((i) => [i, { name: (col("people")[i] || {}).name || (i === me.id ? me.name : "") }])),
  };

  // ---------- saving files on the phone ----------
  const downloadsApi = { save: async ({ filename, data }) => {
    const type = /\.csv$/i.test(filename) ? "text/csv" : /\.pdf$/i.test(filename) ? "application/pdf" : "application/octet-stream";
    const file = new File([data], filename, { type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: filename }); return { status: "saved" }; }
      catch (e) { if (e && e.name === "AbortError") { const err = new Error("declined"); err.code = "declined"; throw err; } }
    }
    const a = document.createElement("a"); a.href = URL.createObjectURL(file); a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    return { status: "saved" };
  } };

  window.claude = { use: async (name) => { await ready; return name === "db" ? dbApi : name === "user" ? userApi : name === "downloads" ? downloadsApi : null; } };

  // ---------- Microsoft sign-in ----------
  const SCOPES = ["Files.ReadWrite", "User.Read"];
  let msalApp = null, account = null;
  const configured = () => !!cfg().clientId && !!window.msal;
  async function initAuth() {
    if (!configured()) return;
    msalApp = new window.msal.PublicClientApplication({
      auth: { clientId: cfg().clientId, authority: cfg().authority || "https://login.microsoftonline.com/consumers", redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: "localStorage" },
    });
    await msalApp.initialize();
    const res = await msalApp.handleRedirectPromise().catch((e) => { setStatus("Sign-in failed: " + (e.errorMessage || e.message || e), "bad"); return null; });
    account = (res && res.account) || msalApp.getAllAccounts()[0] || null;
    if (account) {
      const claims = account.idTokenClaims || {};
      me.id = "ms_" + (claims.oid || account.localAccountId || account.homeAccountId);
      me.name = account.name || account.username || "";
      lsSet("ssi-me", { id: me.id, name: me.name });
    }
  }
  const signIn = () => msalApp && msalApp.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  const signOut = () => msalApp && account && msalApp.logoutRedirect({ account });
  async function token() {
    if (!msalApp || !account) return null;
    try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e && /interaction_required|login_required|consent_required|no_tokens_found|monitor_window_timeout/i.test((e.errorCode || "") + (e.message || ""))) { needLogin = true; return null; }
      throw e;
    }
  }
  let needLogin = false;

  // ---------- OneDrive (Microsoft Graph) ----------
  const G = "https://graph.microsoft.com/v1.0";
  const enc = (p) => String(p).replace(/^\/+|\/+$/g, "").split("/").map(encodeURIComponent).join("/");
  const itemUrl = (p) => G + "/me/drive/root:/" + enc(p);
  async function g(method, url, body, ctype, tok) {
    const h = { Authorization: "Bearer " + tok };
    if (ctype) h["Content-Type"] = ctype;
    const r = await fetch(url, { method, headers: h, body });
    if (r.status === 404) return null;
    if (r.status === 401) { needLogin = true; throw new Error("OneDrive sign-in expired"); }
    if (!r.ok) throw new Error("OneDrive " + method + " " + r.status);
    return r.status === 204 ? {} : (ctype === null ? r : r.json().catch(() => ({})));
  }
  async function listFolder(path, tok) {
    let url = itemUrl(path) + ":/children?$top=999&select=name,eTag,size,lastModifiedDateTime,file,@microsoft.graph.downloadUrl", out = [];
    while (url) { const j = await g("GET", url, null, "", tok); if (!j) return out; out = out.concat(j.value || []); url = j["@odata.nextLink"]; }
    return out;
  }
  async function download(item, tok) {
    const u = item["@microsoft.graph.downloadUrl"];
    const r = u ? await fetch(u) : await fetch(itemUrl(item.path) + ":/content", { headers: { Authorization: "Bearer " + tok } });
    if (!r.ok) throw new Error("Download failed " + r.status);
    return r;
  }

  // ---------- sync ----------
  let syncing = false, again = false, kickT = null, lastErr = "";
  function kick(ms) { clearTimeout(kickT); kickT = setTimeout(() => sync("change"), ms == null ? 2500 : ms); }
  const device = (() => { let d = lsGet("ssi-device", ""); if (!d) { d = Math.random().toString(36).slice(2, 8); lsSet("ssi-device", d); } return d; })();

  async function sync(why) {
    if (syncing) { again = true; return; }
    if (!configured()) { setStatus("Saved on this phone · send entries with Send to Excel · get Excel data with Update from Excel", ""); return; }
    if (!account || needLogin) { setStatus(needLogin ? "Sign in again to sync with OneDrive" : "Sign in to sync with OneDrive", "warn", "signin"); return; }
    if (!navigator.onLine) { setStatus("Offline · " + waitingText() + "saved on this phone", "warn"); return; }
    syncing = true; again = false; setStatus("Syncing…", "");
    try {
      const tok = await token(); if (!tok) { syncing = false; return sync(why); }
      const c = cfg(), app = c.entriesFolder + "/app";
      // 1. send this phone's changes
      for (const k of Object.keys(dirty())) {
        const [cn, id] = k.split("/"), v = col(cn)[id];
        if (v === undefined) await g("DELETE", itemUrl(app + "/" + cn + "/" + id + ".json"), null, "", tok);
        else await g("PUT", itemUrl(app + "/" + cn + "/" + id + ".json") + ":/content", JSON.stringify(v), "application/json", tok);
        const d = Object.assign({}, dirty()); delete d[k]; await setMeta("dirty", d);
      }
      if (me.id !== "local" && (col("people")[me.id] || {}).name !== me.name) await write("people", me.id, { name: me.name, _at: new Date().toISOString() });
      // 2. get other phones' changes
      const tags = meta.etags || {};
      for (const cn of SYNCED) {
        const items = await listFolder(app + "/" + cn, tok);
        for (const it of items) {
          if (!it.file || !/\.json$/.test(it.name)) continue;
          const id = it.name.slice(0, -5), key = cn + "/" + id;
          if (tags[key] === it.eTag || dirty()[key]) continue;
          const v = await (await download(it, tok)).json();
          const cur = col(cn)[id];
          if (!cur || String(v._at || "") >= String(cur._at || "")) await write(cn, id, v, true);
          tags[key] = it.eTag;
        }
      }
      await setMeta("etags", tags);
      // 3. fresh data from Excel
      const H = window.__ssiHooks;
      if (H && H.ready()) {
        const [xi, di] = await Promise.all([g("GET", itemUrl(c.dataFolder + "/" + c.xlsm), null, "", tok), g("GET", itemUrl(c.dataFolder + "/" + c.db), null, "", tok)]);
        if (!xi || !di) setStatus("Excel files not found in OneDrive/" + c.dataFolder, "bad");
        else if (xi.eTag + "|" + di.eTag !== meta.dataTag) {
          setStatus("Reading Excel data from OneDrive…", "");
          const [xb, dbb] = await Promise.all([(await download(xi, tok)).arrayBuffer(), (await download(di, tok)).arrayBuffer()]);
          const summary = await H.autoUpdate(xb, dbb);
          await setMeta("dataTag", xi.eTag + "|" + di.eTag);
          await setMeta("dataAt", new Date().toISOString());
          await setMeta("dataSummary", summary);
        }
        // 4. pending entries as a file for Excel's Import from phone
        const nP = H.pending(), name = c.entriesFolder + "/phone-entries-" + device + ".csv";
        if (nP > 0) {
          const text = await H.exportCsv();
          // put it back if Excel's Import from phone moved it to Done but some entries are still waiting
          const there = text === meta.lastCsv ? await g("GET", itemUrl(name), null, "", tok) : null;
          if (text !== meta.lastCsv || !there) { await g("PUT", itemUrl(name) + ":/content", text, "text/csv", tok); await setMeta("lastCsv", text); }
        } else if (meta.lastCsv) { await g("DELETE", itemUrl(name), null, "", tok); await setMeta("lastCsv", undefined); }
      }
      await setMeta("syncAt", new Date().toISOString());
      lastErr = "";
      setStatus("Synced " + hhmm(meta.syncAt) + " · " + waitingText() + "Excel data " + (meta.dataAt ? hhmm(meta.dataAt) : "not loaded yet"), "ok");
    } catch (e) {
      lastErr = e && e.message ? e.message : String(e);
      setStatus("Sync problem: " + lastErr + " · entries are safe on this phone", "bad", needLogin ? "signin" : "retry");
    }
    syncing = false;
    if (again) kick(500);
  }
  const waitingText = () => { const k = Object.keys(dirty()).length; return k ? k + " change" + (k > 1 ? "s" : "") + " waiting · " : ""; };
  const hhmm = (iso) => { const d = new Date(iso); return isNaN(d) ? "" : d.toLocaleString("en-IN", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" }); };

  // ---------- small status bar + settings ----------
  let bar = null;
  function setStatus(text, tone, action) {
    if (!bar) return;
    bar.className = "odbar " + (tone || "");
    bar.querySelector(".odt").textContent = text;
    const b = bar.querySelector(".oda");
    b.hidden = (syncing && !action) || (!configured() && !action);
    b.textContent = action === "signin" ? "Sign in" : action === "retry" ? "Try again" : configured() ? (account ? "Sync now" : "Sign in") : "Set up";
    b.dataset.act = action || (configured() ? (account ? "sync" : "signin") : "settings");
  }
  function ui() {
    const css = document.createElement("style");
    css.textContent = `.odbar{display:flex;gap:8px;align-items:center;justify-content:space-between;font-size:12px;line-height:1.3;padding:6px 0 0;margin-top:6px;border-top:1px solid rgba(255,255,255,.18);color:rgba(255,255,255,.85)}
.odbar.ok .odt::before{content:"● ";color:#7ee08a}.odbar.warn .odt::before{content:"● ";color:#b26a00}.odbar.bad .odt::before{content:"● ";color:#c62828}
.odbar .odb{display:flex;gap:6px;flex:none}.odbar button{font:inherit;font-weight:600;border:1px solid rgba(255,255,255,.4);background:transparent;color:#fff;border-radius:8px;padding:4px 10px}
.odset{position:fixed;inset:0;background:rgba(0,0,0,.35);display:flex;align-items:flex-end;justify-content:center;z-index:9999}.odset form{background:var(--bg,#fff);color:var(--text,#123);width:100%;max-width:520px;border-radius:14px 14px 0 0;padding:16px;display:grid;gap:10px;font-size:14px}
.odset label{display:grid;gap:4px}.odset input{font:inherit;padding:8px;border:1px solid var(--line,#c9d3d9);border-radius:8px}.odset .row2{display:flex;gap:8px;justify-content:flex-end}
.odup{position:fixed;left:12px;right:12px;bottom:calc(12px + env(safe-area-inset-bottom,0px));z-index:9998;display:flex;gap:10px;align-items:center;justify-content:space-between;background:#0f3b4c;color:#fff;border-radius:12px;padding:10px 14px;font-size:14px}.odup button{font:inherit;font-weight:700;border:0;border-radius:8px;padding:6px 12px}`;
    document.head.appendChild(css);
    bar = document.createElement("div"); bar.className = "odbar"; bar.setAttribute("role", "status");
    bar.innerHTML = `<span class="odt">Starting…</span><span class="odb"><button class="oda" type="button"></button><button class="ods" type="button" aria-label="OneDrive settings">⚙</button></span>`;
    const head = document.querySelector("header") || document.body.firstElementChild;
    head ? head.appendChild(bar) : document.body.prepend(bar);
    bar.querySelector(".oda").onclick = (e) => { const a = e.target.dataset.act; if (a === "signin") signIn(); else if (a === "settings") settings(); else sync("button"); };
    bar.querySelector(".ods").onclick = settings;
  }
  function settings() {
    const c = cfg(), w = document.createElement("div"); w.className = "odset";
    const f = (k, label, hint) => `<label>${label}<input name="${k}" value="${String(c[k] || "").replace(/"/g, "&quot;")}"${hint ? ` placeholder="${hint}"` : ""}></label>`;
    w.innerHTML = `<form><span style="font-size:12px;color:var(--muted,#567)">${window.__ssiVersion ? window.__ssiVersion().replace(/</g, "&lt;") : ""}</span><strong>OneDrive</strong>
      ${account ? `<span>Signed in as ${(me.name || "").replace(/</g, "&lt;")}</span>` : ""}
      ${f("userName", "Your name (shown on entries)")}
      ${f("dataFolder", "Folder with the Excel workbook and database", "SSI/PRODUCTION/Leather/PROCESSES/2025-26")}
      ${f("xlsm", "Workbook file name")}${f("db", "Database file name")}
      ${f("entriesFolder", "Phone entries folder (Excel imports from here)", "SSI/PRODUCTION/Phone entries")}
      <details><summary>Advanced</summary>${f("clientId", "App (client) ID")}${f("authority", "Sign-in authority")}</details>
      <span class="row2">${account ? `<button type="button" data-x="out">Sign out</button>` : ""}<button type="button" data-x="cancel">Cancel</button><button type="submit">Save</button></span></form>`;
    document.body.appendChild(w);
    w.onclick = (e) => { const x = e.target.dataset && e.target.dataset.x; if (e.target === w || x === "cancel") w.remove(); if (x === "out") signOut(); };
    w.querySelector("form").onsubmit = (e) => {
      e.preventDefault(); const o = {};
      new FormData(e.target).forEach((v, k) => { v = String(v).trim(); if (v !== String(CFG0[k] || "")) o[k] = v; });
      lsSet(LSK, o); w.remove(); setMeta("dataTag", undefined).then(() => location.reload());
    };
  }

  // ---------- app updates ----------
  function updates() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.register("sw.js").then((reg) => {
      const offer = (wk) => {
        if (document.querySelector(".odup")) return;
        const d = document.createElement("div"); d.className = "odup"; d.setAttribute("role", "status");
        d.innerHTML = `<span>A new version of the app is ready.</span><button type="button">Update</button>`;
        d.querySelector("button").onclick = () => { wk.postMessage("skip"); d.querySelector("button").disabled = true; };
        document.body.appendChild(d);
      };
      if (reg.waiting && navigator.serviceWorker.controller) offer(reg.waiting);
      reg.addEventListener("updatefound", () => { const wk = reg.installing; wk && wk.addEventListener("statechange", () => { if (wk.state === "installed" && navigator.serviceWorker.controller) offer(wk); }); });
      setInterval(() => reg.update().catch(() => {}), 30 * 60 * 1000);
    }).catch(() => {});
    let reloaded = false;
    navigator.serviceWorker.addEventListener("controllerchange", () => { if (!reloaded) { reloaded = true; location.reload(); } });
  }

  // ---------- start ----------
  window.__ssiSync = { sync: () => sync("manual"), state: () => ({ meta, dirty: dirty(), account: !!account, me }) };
  document.addEventListener("DOMContentLoaded", async () => {
    ui(); updates();
    await ready;
    try { await initAuth(); } catch (e) { setStatus("Sign-in could not start: " + (e.message || e), "bad"); }
    const waitApp = () => new Promise((ok) => { const t0 = Date.now(), tick = () => (window.__ssiHooks && window.__ssiHooks.ready()) || Date.now() - t0 > 15000 ? ok() : setTimeout(tick, 200); tick(); });
    await waitApp();
    sync("start");
    window.addEventListener("online", () => kick(300));
    window.addEventListener("offline", () => sync("offline"));
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") kick(300); });
    setInterval(() => { if (document.visibilityState === "visible") sync("timer"); }, 2 * 60 * 1000);
  });
})();
