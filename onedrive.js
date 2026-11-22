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
  const SYNCED = ["entries", "audits", "people", "invoices", "invset", "quotes", "payments"];

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
    put: async (c, id, v) => write(c, id, Object.assign({ _at: new Date().toISOString(), _by: me.id }, JSON.parse(JSON.stringify(v)))),
  };

  // ---------- who is using it ----------
  const me = { id: lsGet("ssi-me", { id: "local" }).id, name: lsGet("ssi-me", { name: "" }).name || cfg().userName || "" };
  const userApi = {
    id: async () => me.id, me: async () => ({ id: me.id, name: me.name }),
    can: async () => true, canEdit: async () => true, isOwner: () => true,
    profiles: async (ids) => Object.fromEntries([].concat(ids).map((i) => [i, { name: (col("people")[i] || {}).name || (i === me.id ? me.name : "") }])),
  };

  // ---------- saving files on the phone ----------
  // Share sheet first (so the file can go straight to OneDrive). Android refuses some types (.xml), so the same file is
  // offered as .txt (Excel's Invoices > Tally reads tally-sales-*.xml.txt too). If the phone says the tap is too old to share,
  // a "Share" button is shown and the user's tap opens the share sheet. Download is the last resort.
  const declined = () => { const err = new Error("declined"); err.code = "declined"; return err; };
  const tryShare = async (file) => { if (!(navigator.canShare && navigator.canShare({ files: [file] }))) return "no"; try { await navigator.share({ files: [file], title: file.name }); return "ok"; } catch (e) { return e && e.name === "AbortError" ? "abort" : e && e.name === "NotAllowedError" ? "tap" : "no"; } };
  const tapToShare = (file) => new Promise((res) => {
    const box = document.createElement("div");
    box.setAttribute("style", "position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.45);display:flex;align-items:flex-end;justify-content:center;padding:16px");
    box.innerHTML = '<div style="background:var(--surface,#fff);color:var(--ink,#111);border-radius:14px;padding:16px;max-width:420px;width:100%;font:15px system-ui"><b style="display:block;margin-bottom:6px">File ready</b><span style="display:block;word-break:break-all;margin-bottom:12px">' + file.name.replace(/[<>&]/g, "") + '</span><button data-a="share" style="width:100%;padding:12px;border:0;border-radius:10px;background:var(--accent,#1b6ef3);color:var(--accent-ink,#fff);font-weight:600;font-size:16px">Share to OneDrive…</button><button data-a="dl" style="width:100%;margin-top:8px;padding:10px;border:1px solid var(--line,#ccc);border-radius:10px;background:transparent;color:inherit">Download instead</button><button data-a="x" style="width:100%;margin-top:8px;padding:10px;border:0;background:transparent;color:inherit">Cancel</button></div>';
    document.body.appendChild(box);
    box.addEventListener("click", async (e) => { const a = (e.target.closest("button") || {}).dataset; if (!a || !a.a) return;
      if (a.a === "share") { const r = await tryShare(file); if (r === "ok") { box.remove(); res("ok"); } else if (r === "abort") { /* stay open */ } else { box.remove(); res("dl"); } }
      else { box.remove(); res(a.a); } });
  });
  const dlFile = (file) => { const a = document.createElement("a"); a.href = URL.createObjectURL(file); a.download = file.name; document.body.appendChild(a); a.click(); a.remove(); };
  // a PDF goes straight into the PDF folder in OneDrive (settings); then the phone can still share it (WhatsApp, mail)
  const pdfBox = (file, said, ok) => new Promise((res) => {
    const box = document.createElement("div"), e = (t) => String(t).replace(/[<>&]/g, "");
    box.setAttribute("style", "position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.45);display:flex;align-items:flex-end;justify-content:center;padding:16px");
    box.innerHTML = '<div class="pdfbox" style="background:var(--surface,#fff);color:var(--ink,#111);border-radius:14px;padding:16px;max-width:420px;width:100%;font:15px system-ui"><b style="display:block;margin-bottom:6px">' + (ok ? "Saved to OneDrive" : "Not saved to OneDrive") + '</b><span style="display:block;word-break:break-all;margin-bottom:12px;' + (ok ? "" : "color:#b45309") + '">' + e(said) + '</span><button data-a="share" style="width:100%;padding:12px;border:0;border-radius:10px;background:' + (ok ? "transparent;border:1px solid var(--line,#ccc);color:inherit" : "var(--accent,#1b6ef3);color:var(--accent-ink,#fff)") + ';font-weight:600;font-size:16px">Share…</button>' + (ok ? "" : '<button data-a="dl" style="width:100%;margin-top:8px;padding:10px;border:1px solid var(--line,#ccc);border-radius:10px;background:transparent;color:inherit">Download</button>') + '<button data-a="x" style="width:100%;margin-top:8px;padding:12px;border:0;border-radius:10px;' + (ok ? "background:var(--accent,#1b6ef3);color:var(--accent-ink,#fff);font-weight:600" : "background:transparent;color:inherit") + '">Done</button></div>';
    document.body.appendChild(box);
    box.addEventListener("click", async (ev) => { const a = (ev.target.closest("button") || {}).dataset; if (!a || !a.a) return;
      if (a.a === "share") { const r = await tryShare(file); if (r === "abort") return; if (r !== "ok") dlFile(file); }
      else if (a.a === "dl") dlFile(file);
      box.remove(); res(a.a); });
  });
  const pdfFolder = () => String(cfg().pdfFolder || cfg().dataFolder || "").replace(/^\/+|\/+$/g, "");
  const TALLY_DIR0 = "SSI/Accountancy/tally data/SSI/XML DATA";
  const tallyFolder = () => String(cfg().tallyFolder || TALLY_DIR0).replace(/^\/+|\/+$/g, "");
  const invFolder = () => String(cfg().invoiceFolder || "").replace(/^\/+|\/+$/g, "") || pdfFolder();
  async function pdfToOneDrive(file, kind) {
    const folder = kind === "invoice" ? invFolder() : pdfFolder(); if (!folder) return { ok: false, said: "No PDF folder is set (OneDrive settings)." };
    if (!account) return { ok: false, said: "Not signed in to OneDrive." };
    if (!navigator.onLine) return { ok: false, said: "No internet just now." };
    try { const tok = await token(); if (!tok) return { ok: false, said: "Sign in to OneDrive again." };
      await upload(folder + "/" + file.name.replace(/[\\/:*?"<>|#%]+/g, "-"), await file.arrayBuffer(), tok, "application/pdf");
      return { ok: true, said: folder.replace(/\//g, " › ") + " › " + file.name };
    } catch (e) { return { ok: false, said: "Could not upload: " + (e.message || e) }; }
  }
  // a PDF first shows a box: the name can be changed, then Save to OneDrive and / or Share (WhatsApp, mail) or Download
  const pdfAsk = (data, filename, kind) => new Promise((res) => {
    const e = (t) => String(t).replace(/[<>&"]/g, ""), folder = kind === "invoice" ? invFolder() : pdfFolder(), base = filename.replace(/\.pdf$/i, "");
    const box = document.createElement("div"); let saved = false;
    box.setAttribute("style", "position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.45);display:flex;align-items:flex-end;justify-content:center;padding:16px");
    const btn = (a, t, main) => '<button data-a="' + a + '" style="width:100%;margin-top:8px;padding:12px;border-radius:10px;font-size:16px;' + (main ? "border:0;background:var(--accent,#1b6ef3);color:var(--accent-ink,#fff);font-weight:600" : "border:1px solid var(--line,#ccc);background:transparent;color:inherit") + '">' + t + "</button>";
    box.innerHTML = '<div class="pdfbox" style="background:var(--surface,#fff);color:var(--ink,#111);border-radius:14px;padding:16px;max-width:440px;width:100%;font:15px system-ui"><b style="display:block;margin-bottom:8px">PDF ready</b>' +
      '<label style="display:block;font-size:13px;color:var(--muted,#666)">File name<input id="pdfname" value="' + e(base) + '" autocomplete="off" style="display:block;width:100%;box-sizing:border-box;margin-top:4px;padding:10px;font-size:16px;border:1px solid var(--line,#ccc);border-radius:8px;background:var(--bg,#fff);color:inherit"></label>' +
      '<span id="pdfsaid" style="display:block;font-size:13px;margin:8px 0 2px;color:var(--muted,#666);word-break:break-all">' + (folder ? "OneDrive folder: " + e(folder.replace(/\//g, " › ")) : "No OneDrive PDF folder is set (OneDrive settings).") + "</span>" +
      btn("pv", "Preview") + btn("od", "Save to OneDrive", true) + btn("share", "Share PDF…") + btn("dl", "Download") + '<button data-a="x" style="width:100%;margin-top:8px;padding:10px;border:0;background:transparent;color:inherit">Close</button></div>';
    document.body.appendChild(box);
    const said = box.querySelector("#pdfsaid"), nameIn = box.querySelector("#pdfname");
    const file = () => { const n = (nameIn.value.trim() || base).replace(/[\\/:*?"<>|#%]+/g, "-").replace(/\.pdf$/i, ""); return new File([data], n + ".pdf", { type: "application/pdf" }); };
    box.addEventListener("click", async (ev) => { const b = ev.target.closest("button"), a = b && b.dataset.a; if (!a) return;
      if (a === "x") { box.remove(); res(saved); return; }
      if (a === "dl") { dlFile(file()); return; }
      if (a === "pv") { if (window.__ssiPreview) window.__ssiPreview(data, file().name, []); return; } // look at it without saving (22-Nov-2026)
      if (a === "share") { const r = await tryShare(file()); if (r !== "ok" && r !== "abort") dlFile(file()); return; }
      if (a === "od") { b.disabled = true; said.style.color = "var(--muted,#666)"; said.textContent = "Saving…"; const r = await pdfToOneDrive(file(), kind); b.disabled = false;
        said.textContent = r.ok ? "Saved: " + r.said : r.said; said.style.color = r.ok ? "var(--ok,#1d7a46)" : "#b45309"; if (r.ok) { saved = true; b.textContent = "Saved ✓ (save again)"; } } });
  });
  // a Tally invoice file goes straight into Excel's XML DATA folder as .xml.txt; Excel's automatic import enters it in Tally
  const tallyBox = (file, said, ok) => new Promise((res) => {
    const box = document.createElement("div"), e = (t) => String(t).replace(/[<>&]/g, "");
    box.setAttribute("style", "position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.45);display:flex;align-items:flex-end;justify-content:center;padding:16px");
    box.innerHTML = '<div class="pdfbox" style="background:var(--surface,#fff);color:var(--ink,#111);border-radius:14px;padding:16px;max-width:440px;width:100%;font:15px system-ui"><b style="display:block;margin-bottom:6px">' + (ok ? "Sent to Excel for Tally" : "Not saved to OneDrive") + '</b><span style="display:block;word-break:break-all;margin-bottom:8px;' + (ok ? "" : "color:#b45309") + '">' + e(said) + "</span>" + (ok ? '<span style="display:block;font-size:13px;color:var(--muted,#666);margin-bottom:8px">Excel enters it in Tally within about a minute while Excel and Tally are open on the PC. The invoice shows In Tally after that.</span>' : "") +
      '<button data-a="share" style="width:100%;margin-top:4px;padding:12px;border-radius:10px;font-size:16px;border:1px solid var(--line,#ccc);background:transparent;color:inherit">Share the file…</button><button data-a="x" style="width:100%;margin-top:8px;padding:12px;border:0;border-radius:10px;background:var(--accent,#1b6ef3);color:var(--accent-ink,#fff);font-weight:600;font-size:16px">Done</button></div>';
    document.body.appendChild(box);
    box.addEventListener("click", async (ev) => { const b = ev.target.closest("button"), a = b && b.dataset.a; if (!a) return;
      if (a === "share") { const r = await tryShare(file); if (r !== "ok" && r !== "abort") dlFile(file); return; }
      box.remove(); res(ok); });
  });
  async function tallyToOneDrive(file) {
    if (!account) return { ok: false, said: "Not signed in to OneDrive." };
    if (!navigator.onLine) return { ok: false, said: "No internet just now." };
    try { const tok = await token(); if (!tok) return { ok: false, said: "Sign in to OneDrive again." };
      await upload(tallyFolder() + "/" + file.name.replace(/[\\/:*?"<>|#%]+/g, "-"), await file.arrayBuffer(), tok, "text/plain");
      return { ok: true, said: tallyFolder().replace(/\//g, " › ") + " › " + file.name };
    } catch (e) { return { ok: false, said: "Could not upload: " + (e.message || e) }; }
  }
  const downloadsApi = { save: async ({ filename, data, kind }) => {
    if (/^tally-sales-.*\.xml$/i.test(filename)) {
      const file = new File([data], filename + ".txt", { type: "text/plain" }), r = await tallyToOneDrive(file);
      await tallyBox(file, r.said, r.ok); return { status: "saved", onedrive: r.ok };
    }
    if (/\.pdf$/i.test(filename) && cfg().pdfAuto !== "off") { const r = await pdfAsk(data, filename, kind); return { status: "saved", onedrive: r }; }
    const type = /\.csv$/i.test(filename) ? "text/csv" : /\.pdf$/i.test(filename) ? "application/pdf" : /\.xml$/i.test(filename) ? "text/xml" : /\.json$/i.test(filename) ? "application/json" : "application/octet-stream";
    let file = new File([data], filename, { type });
    let r = await tryShare(file);
    if (r === "no" && /\.(xml|json)$/i.test(filename)) { file = new File([data], filename + ".txt", { type: "text/plain" }); r = await tryShare(file); }
    if (r === "abort") throw declined();
    if (r === "tap") { r = await tapToShare(file); if (r === "x") throw declined(); }
    if (r !== "ok") dlFile(file);
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
      auth: { clientId: cfg().clientId, authority: cfg().authority || "https://login.microsoftonline.com/common", redirectUri: location.origin + location.pathname },
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

  // ---------- backups in OneDrive ----------
  const slug = (t) => String(t || "").normalize("NFKD").replace(/[^\w]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 30) || "phone";
  const snapName = (day) => day + "_" + device + "_" + slug(me.name) + ".json";
  const snapParse = (n) => { const m = /^(\d{4}-\d{2}-\d{2})_([a-z0-9]+)_(.*)\.json$/i.exec(n); return m ? { day: m[1], device: m[2], who: m[3].replace(/-/g, " ") } : null; };
  // files over 4 MB need an upload session; small ones go in one PUT
  async function upload(path, text, tok, type) {
    const blob = new Blob([text], { type: type || "application/json" });
    if (blob.size < 3.9e6) return g("PUT", itemUrl(path) + ":/content", blob, type || "application/json", tok);
    const ses = await g("POST", itemUrl(path) + ":/createUploadSession", JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }), "application/json", tok);
    const step = 320 * 1024 * 10;
    for (let i = 0; i < blob.size; i += step) {
      const part = blob.slice(i, Math.min(blob.size, i + step));
      const r = await fetch(ses.uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${i}-${i + part.size - 1}/${blob.size}` }, body: part });
      if (!r.ok) throw new Error("OneDrive upload " + r.status);
    }
  }
  // ---------- who keeps the records ----------
  // <entries folder>/app/records.json says it for every phone: {mode: "app" | "excel", base, since, by}. In app mode every phone
  // starts from the same base file (app/base/<name>.sqlite, Excel's data when the app took over) and writes the entries itself;
  // Excel's files only bring Tally's chemical stock and prices. Giving the records back lets Excel import every entry since.
  const recLocal = (on) => { try { if (on) localStorage.setItem("ssi-records", "app"); else localStorage.removeItem("ssi-records"); } catch (e) {} };
  const recIsOn = () => { try { return localStorage.getItem("ssi-records") === "app"; } catch (e) { return false; } };
  async function readRecords(tok) {
    const it = await g("GET", itemUrl(cfg().entriesFolder + "/app/records.json"), null, "", tok);
    if (!it) return null;
    if (it.eTag === meta.recTag && meta.records) return meta.records;
    const r = await (await download(it, tok)).json(); await setMeta("records", r); await setMeta("recTag", it.eTag); return r;
  }
  async function records(to) {
    const H = window.__ssiHooks, tok = await token(); if (!tok) throw new Error("Sign in to OneDrive first.");
    if (!H || !H.ready()) throw new Error("The app is still loading. Try again in a moment.");
    const app = cfg().entriesFolder + "/app", by = me.name || me.id, since = new Date().toISOString();
    if (to === "app") {
      await sync("records"); while (syncing) await new Promise((r) => setTimeout(r, 300));
      // entries still waiting for Excel are simply written by the app from now on (Excel never marks them imported in app mode)
      const bytes = H.baseBytes(); if (!bytes) throw new Error("Excel's data is not loaded yet.");
      const name = "base-" + since.replace(/[-:]/g, "").slice(0, 15) + ".sqlite";
      await upload(app + "/base/" + name, bytes, tok, "application/octet-stream");
      await g("PUT", itemUrl(app + "/records.json") + ":/content", JSON.stringify({ mode: "app", base: name, since, by }), "application/json", tok);
      await setMeta("baseName", name); await setMeta("recTag", undefined); recLocal(true);
      await H.useBase(bytes); setStatus("The app keeps the records", "ok");
      return "Done. From now on the app keeps the records. Other phones switch at their next sync.";
    }
    await g("PUT", itemUrl(app + "/records.json") + ":/content", JSON.stringify({ mode: "excel", since, by }), "application/json", tok);
    recLocal(false); await setMeta("baseName", undefined); await setMeta("dataTag", undefined); await setMeta("recTag", undefined);
    kick(300);
    return "Done. Excel keeps the records again and will import the entries made since the app took over.";
  }
  window.__ssiOD = {
    records,
    settings: () => settings(),
    ready: () => configured() && !!account && !needLogin,
    syncNow: async () => { await sync("restore"); while (syncing) await new Promise((r) => setTimeout(r, 300)); },
    // "Sync with Excel and Tally now": this phone's entries go up first, then excel-now.txt asks Excel to run its import at once (21-Nov-2026)
    askExcel: async () => {
      if (!account) throw new Error("Not signed in to OneDrive.");
      if (!navigator.onLine) throw new Error("No internet just now.");
      // only the small request file is waited for; this phone's entries go up in the background straight after (Excel looks once a minute)
      const tok = await token(); if (!tok) throw new Error("Sign in to OneDrive again.");
      await g("PUT", itemUrl(cfg().entriesFolder + "/excel-now.txt") + ":/content", new Date().toISOString() + " " + device + " " + (me.name || ""), "text/plain", tok);
      kick(50); [75, 120, 180, 270].forEach((s2) => setTimeout(() => sync("change"), s2 * 1000));
      // Excel's answer and Tally's sales are two small files: look every 10 s for 4 minutes, without waiting for Excel's whole data
      const t0 = Date.now(), poll = async () => { try { const t = await token(); if (t) await lightCheck(t); } catch (e) { /* next time */ } if (Date.now() - t0 < 240000 && !lightDone) setTimeout(poll, 10000); };
      lightDone = false; askedAt = new Date().toISOString(); setTimeout(poll, 15000);
      return true;
    },
    list: async () => {
      const tok = await token(); if (!tok) throw new Error("Sign in to OneDrive again.");
      const items = await listFolder(cfg().entriesFolder + "/app/backups", tok);
      return items.filter((it) => it.file && snapParse(it.name)).map((it) => Object.assign(snapParse(it.name), { name: it.name, at: it.lastModifiedDateTime, size: it.size, mine: snapParse(it.name).device === device }))
        .sort((a, b) => b.name.slice(0, 10).localeCompare(a.name.slice(0, 10)) || String(b.at).localeCompare(String(a.at)));
    },
    get: async (name) => {
      const tok = await token(); if (!tok) throw new Error("Sign in to OneDrive again.");
      const it = await g("GET", itemUrl(cfg().entriesFolder + "/app/backups/" + name), null, "", tok);
      if (!it) throw new Error("That backup is no longer in OneDrive.");
      return (await download(it, tok)).json();
    },
  };

  // ---------- small files from Excel (21-Nov-2026) ----------
  // app/tally-sales.json (Tally's sale vouchers) and app/excel-now-done.json (what Excel did when asked); tags kept in memory so a restart reads them again
  let tsTag = "", doneTag = "", lightDone = true, askedAt = "";
  async function lightCheck(tok) {
    const H = window.__ssiHooks; if (!H || !H.ready() || !H.tallySales) return;
    const app = cfg().entriesFolder + "/app", rd = async (it) => JSON.parse((await (await download(it, tok)).text()).replace(/^\uFEFF/, ""));
    const [ts, dn] = await Promise.all([g("GET", itemUrl(app + "/tally-sales.json"), null, "", tok).catch(() => null), g("GET", itemUrl(app + "/excel-now-done.json"), null, "", tok).catch(() => null)]);
    if (ts && ts.eTag !== tsTag) { H.tallySales(await rd(ts)); tsTag = ts.eTag; }
    if (dn && dn.eTag !== doneTag) { const o = await rd(dn); doneTag = dn.eTag; if (H.excelDone) H.excelDone(o); if (askedAt && new Date(String(o.at).replace(" ", "T")) >= new Date(new Date(askedAt).getTime() - 60000)) lightDone = true; }
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
      // 3. who keeps the records, then fresh data from Excel
      const H = window.__ssiHooks;
      if (H && H.ready()) {
        try { await lightCheck(tok); } catch (e) { /* not there yet */ }
        const rec = await readRecords(tok), appMode = !!rec && rec.mode === "app";
        if (appMode && rec.base && rec.base !== meta.baseName) {
          setStatus("Getting the app's records from OneDrive…", "");
          const bi = await g("GET", itemUrl(app + "/base/" + rec.base), null, "", tok);
          if (!bi) throw new Error("The records file " + rec.base + " is missing in OneDrive");
          recLocal(true); await H.useBase(new Uint8Array(await (await download(bi, tok)).arrayBuffer()));
          await setMeta("baseName", rec.base); await setMeta("dataTag", undefined);
        } else if (!appMode && recIsOn()) { recLocal(false); await setMeta("baseName", undefined); await setMeta("dataTag", undefined); }
        // app keeps the records: Tally's figures come from Excel's small file app/tally-data.json; the workbook is never downloaded (22-Nov-2026)
        let tj = null; if (appMode) { try { tj = await g("GET", itemUrl(app + "/tally-data.json"), null, "", tok); } catch (e) { tj = null; } }
        if (tj) {
          if (tj.eTag !== meta.tjTag) { setStatus("Reading Tally stock and prices…", "");
            const o = JSON.parse((await (await download(tj, tok)).text()).replace(/^\uFEFF/, ""));
            const summary = await H.tallyJson(o); await setMeta("tjTag", tj.eTag); await setMeta("dataAt", new Date().toISOString()); await setMeta("dataSummary", summary); }
        } else {
        const [xi, di] = await Promise.all([g("GET", itemUrl(c.dataFolder + "/" + c.xlsm), null, "", tok), g("GET", itemUrl(c.dataFolder + "/" + c.db), null, "", tok)]);
        if (!xi || !di) setStatus("Excel files not found in OneDrive/" + c.dataFolder, "bad");
        else if (xi.eTag + "|" + di.eTag !== meta.dataTag) {
          setStatus(appMode ? "Reading Tally stock and prices from Excel…" : "Reading Excel data from OneDrive…", "");
          const [xb, dbb] = await Promise.all([(await download(xi, tok)).arrayBuffer(), (await download(di, tok)).arrayBuffer()]);
          const summary = appMode ? await H.tallyMerge(xb, dbb) : await H.autoUpdate(xb, dbb);
          await setMeta("dataTag", xi.eTag + "|" + di.eTag);
          await setMeta("dataAt", new Date().toISOString());
          await setMeta("dataSummary", summary);
        }
        }
        // 3b. app mode: once a day, close entries older than 3 days into a new base so phones replay less (at least 50)
        if (appMode && H.compact && Date.now() - Date.parse(rec.compacted_at || rec.since || 0) > 864e5) {
          const r = H.compact(new Date(Date.now() - 3 * 864e5).toISOString(), 50);
          if (r) {
            const now = new Date().toISOString(), name = "base-" + now.replace(/[-:]/g, "").slice(0, 15) + ".sqlite";
            await upload(app + "/base/" + name, r.bytes, tok, "application/octet-stream");
            await setMeta("recTag", undefined); const again = await readRecords(tok);
            if (again && again.mode === "app" && again.base === rec.base) {
              await g("PUT", itemUrl(app + "/records.json") + ":/content", JSON.stringify(Object.assign({}, again, { base: name, compacted_at: now, compacted_by: me.name || me.id, folded: (again.folded || 0) + r.n })), "application/json", tok);
              await setMeta("recTag", undefined); await setMeta("baseName", name); await H.useBase(r.bytes);
              // keep the last three bases in OneDrive
              const olds = (await listFolder(app + "/base", tok)).filter((it) => it.file && /^base-.*\.sqlite$/.test(it.name)).sort((x, y) => y.name.localeCompare(x.name));
              for (const it of olds.slice(3)) await g("DELETE", itemUrl(app + "/base/" + it.name), null, "", tok);
            } else await g("DELETE", itemUrl(app + "/base/" + name), null, "", tok);
          }
        }
        // 4. pending entries as a file for Excel's Import from phone
        // records kept by the app: nothing goes to Excel any more
        let recApp = false; try { recApp = localStorage.getItem("ssi-records") === "app"; } catch (e) { /* default */ }
        // app mode: Excel is the backup copy, so every entry since the switch that Excel hasn't taken goes in the file (app/excel-mirror.json lists what it took)
        let mirror = null;
        if (appMode && H.mirrorBehind) {
          try { const mi = await g("GET", itemUrl(app + "/excel-mirror.json"), null, "", tok);
            if (mi && mi.eTag !== meta.mirTag) { const o = JSON.parse((await (await download(mi, tok)).text()).replace(/^\uFEFF/, "")); await setMeta("mirDone", o.done || []); await setMeta("mirAt", o.at || ""); await setMeta("mirTag", mi.eTag); } } catch (e) { /* Excel hasn't written it yet */ }
          mirror = { since: rec.since || "", done: new Set(meta.mirDone || []) };
          await setMeta("mirBehind", H.mirrorBehind(mirror));
        }
        const nP = mirror ? H.mirrorBehind(mirror).n : H.pending(), name = c.entriesFolder + "/phone-entries-" + device + ".csv";
        if (nP > 0) {
          const text = await H.exportCsv(mirror || undefined);
          // put it back if Excel's Import from phone moved it to Done but some entries are still waiting
          const there = text === meta.lastCsv ? await g("GET", itemUrl(name), null, "", tok) : null;
          if (text !== meta.lastCsv || !there) { await g("PUT", itemUrl(name) + ":/content", text, "text/csv", tok); await setMeta("lastCsv", text); }
        } else if (meta.lastCsv) { await g("DELETE", itemUrl(name), null, "", tok); await setMeta("lastCsv", undefined); }
        // what Excel's automatic import last did (excel-status.json, written by the workbook)
        try { const xs = await g("GET", itemUrl(c.entriesFolder + "/excel-status.json"), null, "", tok);
          if (xs && xs.eTag !== meta.xsTag) { await setMeta("excelStatus", await (await download(xs, tok)).json()); await setMeta("xsTag", xs.eTag); } } catch (e) { /* not set up yet */ }
      }
      // 5. this phone's daily backup: <entries folder>/app/backups/<day>_<phone>_<name>.json, the last 30 kept per phone
      if (H && H.backupBody) {
        const day = new Date().toLocaleDateString("en-CA");
        if (meta.snapDay !== day) {
          const body = Object.assign(await H.backupBody(), { device, who: me.name || "" }), txt = JSON.stringify(body);
          await upload(app + "/backups/" + snapName(day), txt, tok);
          await setMeta("snapDay", day);
          try { localStorage.setItem("ssi-last-backup", body.made_at); } catch (e) {}
          const mine = (await listFolder(app + "/backups", tok)).filter((it) => it.file && snapParse(it.name) && snapParse(it.name).device === device).sort((a, b) => b.name.localeCompare(a.name));
          for (const it of mine.slice(30)) await g("DELETE", itemUrl(app + "/backups/" + it.name), null, "", tok);
        }
      }
      await setMeta("syncAt", new Date().toISOString());
      lastErr = "";
      setStatus("Synced " + hhmm(meta.syncAt) + " · " + waitingText() + (recIsOn() ? "records kept by the app · Tally stock " + (meta.dataAt ? hhmm(meta.dataAt) : "not read yet") + (meta.mirBehind ? " · Excel backup " + (meta.mirBehind.old ? meta.mirBehind.old + " entr" + (meta.mirBehind.old === 1 ? "y" : "ies") + " behind" : "up to date") : "") : "Excel data " + (meta.dataAt ? hhmm(meta.dataAt) : "not loaded yet") + xsText()), "ok");
    } catch (e) {
      lastErr = e && e.message ? e.message : String(e);
      setStatus("Sync problem: " + lastErr + " · entries are safe on this phone", "bad", needLogin ? "signin" : "retry");
    }
    syncing = false;
    if (again) kick(500);
  }
  const xsText = () => { const x = meta.excelStatus; if (!x || !x.at) return ""; const old = Date.now() - new Date(x.at) > Math.max(30, 3 * (x.every || 5)) * 60000;
    return " · Excel " + (x.on === false ? "auto import off" : old ? "not checking (open since " + hhmm(x.at) + "?)" : "checked " + hhmm(x.at)); };
  const waitingText = () => { const k = Object.keys(dirty()).length; return k ? k + " change" + (k > 1 ? "s" : "") + " waiting · " : ""; };
  const hhmm = (iso) => { const d = new Date(iso); return isNaN(d) ? "" : d.toLocaleString("en-IN", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" }); };

  // ---------- small status bar + settings ----------
  let bar = null;
  let lastSt = { text: "Starting…", tone: "" };
  // the last finished sync (good or not) stays visible under "Syncing…" and the other in-progress lines (22-Nov-2026)
  let lastDone = (() => { try { return JSON.parse(localStorage.getItem("ssi-od-lastdone") || "null"); } catch (e) { return null; } })();
  const lastLine = () => (lastSt.tone === "" && lastDone ? "Last sync: " + lastDone.text : "");
  function setStatus(text, tone, action) {
    lastSt = { text, tone: tone || "" };
    if (tone === "ok" || tone === "bad") { lastDone = { text, tone, at: new Date().toISOString() }; try { localStorage.setItem("ssi-od-lastdone", JSON.stringify(lastDone)); } catch (e) {} }
    const ss = document.querySelector(".odset .odst"); if (ss) { ss.className = "odst odbar " + (tone || ""); ss.querySelector(".odt").textContent = text; const l = ss.querySelector(".odl"); if (l) l.textContent = lastLine(); }
    if (!bar) return;
    bar.className = "odbar " + (tone || "");
    bar.querySelector(".odt").textContent = text;
    { const l = bar.querySelector(".odl"); if (l) l.textContent = lastLine(); }
    const b = bar.querySelector(".oda");
    b.hidden = (syncing && !action) || (!configured() && !action);
    b.textContent = action === "signin" ? "Sign in" : action === "retry" ? "Try again" : configured() ? (account ? "Sync now" : "Sign in") : "Set up";
    b.dataset.act = action || (configured() ? (account ? "sync" : "signin") : "settings");
    // OneDrive lives in Menu > Settings > OneDrive (21-Nov-2026); the line under the top bar shows only when the user has to sign in or set up
    bar.hidden = !(b.dataset.act === "signin" || b.dataset.act === "settings");
  }
  function ui() {
    const css = document.createElement("style");
    css.textContent = `.odbar{display:flex;gap:8px;align-items:center;justify-content:space-between;font-size:12px;line-height:1.3;padding:6px 0 0;margin-top:6px;border-top:1px solid var(--line,#d5dde2);color:var(--muted,#5b6b74)}
.odbar.ok .odt::before{content:"● ";color:#2e9d4a}.odbar.warn .odt::before{content:"● ";color:#b26a00}.odbar.bad .odt::before{content:"● ";color:#c62828}
.odbar .odb{display:flex;gap:6px;flex:none}.odl{display:block;font-size:11px;opacity:.8;margin-top:2px}.odl:empty{display:none}
.odbar button{font:inherit;font-weight:600;border:1px solid var(--line,#c9d3d9);background:var(--surface,#fff);color:var(--ink,#15232b);border-radius:8px;padding:4px 10px}
.odset{position:fixed;inset:0;background:rgba(0,0,0,.35);display:flex;align-items:flex-end;justify-content:center;z-index:9999}.odset form{background:var(--surface,#fff);color:var(--ink,#15232b);width:100%;max-width:520px;border-radius:14px 14px 0 0;padding:16px;padding-bottom:calc(16px + env(safe-area-inset-bottom,0px));display:grid;gap:10px;font-size:14px;max-height:calc(100vh - 24px);max-height:calc(100dvh - 24px);overflow-y:auto;overscroll-behavior:contain;-webkit-overflow-scrolling:touch;box-sizing:border-box}.pdfbox{max-height:calc(100vh - 32px);max-height:calc(100dvh - 32px);overflow-y:auto;overscroll-behavior:contain;box-sizing:border-box}
.odset label{display:grid;gap:4px}.odset input{font:inherit;padding:8px;background:var(--bg,#fff);color:var(--ink,#15232b);border:1px solid var(--line,#c9d3d9);border-radius:8px}.odset .row2{display:flex;flex-wrap:wrap;gap:8px;justify-content:flex-end;position:sticky;bottom:-16px;margin:0 -16px -16px;padding:10px 16px calc(10px + env(safe-area-inset-bottom,0px));background:var(--surface,#fff);border-top:1px solid var(--line,#c9d3d9)}
.odup{position:fixed;left:12px;right:12px;bottom:calc(12px + env(safe-area-inset-bottom,0px));z-index:9998;display:flex;gap:10px;align-items:center;justify-content:space-between;background:#0f3b4c;color:#fff;border-radius:12px;padding:10px 14px;font-size:14px}.odup button{font:inherit;font-weight:700;border:0;border-radius:8px;padding:6px 12px}`;
    document.head.appendChild(css);
    bar = document.createElement("div"); bar.className = "odbar"; bar.setAttribute("role", "status");
    bar.innerHTML = `<span><span class="odt">Starting…</span><span class="odl"></span></span><span class="odb"><button class="oda" type="button"></button><button class="ods" type="button" aria-label="OneDrive settings">⚙</button></span>`;
    const head = document.querySelector("header") || document.body.firstElementChild;
    head ? head.appendChild(bar) : document.body.prepend(bar);
    bar.querySelector(".oda").onclick = (e) => { const a = e.target.dataset.act; if (a === "signin") signIn(); else if (a === "settings") settings(); else sync("button"); };
    bar.querySelector(".ods").onclick = settings;
  }
  function settings() {
    const c = cfg(), w = document.createElement("div"); w.className = "odset";
    const f = (k, label, hint) => `<label>${label}<input name="${k}" value="${String(c[k] || "").replace(/"/g, "&quot;")}"${hint ? ` placeholder="${hint}"` : ""}></label>`;
    const act = bar ? bar.querySelector(".oda").dataset.act : "";
    w.innerHTML = `<form><span style="font-size:12px;color:var(--muted,#567)">${window.__ssiVersion ? window.__ssiVersion().replace(/</g, "&lt;") : ""}</span><strong>OneDrive</strong>
      <div class="odst odbar ${lastSt.tone}" style="margin:0;padding:0;border:0"><span><span class="odt">${String(lastSt.text).replace(/</g, "&lt;")}</span><span class="odl">${lastLine().replace(/</g, "&lt;")}</span></span>${configured() ? `<button type="button" data-x="act">${act === "signin" ? "Sign in" : account ? "Sync now" : "Sign in"}</button>` : ""}</div>
      ${account ? `<span>Signed in as ${(me.name || "").replace(/</g, "&lt;")}</span>` : ""}
      ${f("userName", "Your name (shown on entries)")}
      ${f("dataFolder", "Folder with the Excel workbook and database", "SSI/PRODUCTION/Leather/PROCESSES/2025-26")}
      ${f("xlsm", "Workbook file name")}${f("db", "Database file name")}
      ${f("entriesFolder", "Phone entries folder (the app puts entries here; Excel imports them on its own)", "SSI/PRODUCTION/Phone entries")}
      ${f("pdfFolder", "PDF folder (every PDF made in the app is saved here)", c.dataFolder || "SSI/PRODUCTION/Leather/PROCESSES/2025-26")}
      ${f("invoiceFolder", "Invoice PDF folder (empty = the PDF folder above)", "")}
      ${f("tallyFolder", "Tally files folder (Excel's XML DATA; Excel sends files from here to Tally)", TALLY_DIR0)}
      <label style="display:flex;flex-direction:row;align-items:center;gap:8px"><input type="checkbox" name="pdfAutoOn" ${c.pdfAuto === "off" ? "" : "checked"} style="width:auto"> For PDFs, show the PDF box</label>
      <span style="font-size:12px;color:var(--muted,#567);margin-top:-6px">Ticked: every PDF button opens a box where you can <b>Preview</b> the PDF without saving, change its name, save it to OneDrive, share or download it. Unticked: the PDF goes straight to Share, with no box and no preview.</span>
      ${f("syncMinutes", "Check OneDrive every (minutes, while the app is open)", "2")}
      ${(() => { const x = meta.excelStatus; if (!x || !x.at) return `<span style="font-size:12px;color:var(--muted,#567)">Excel's automatic import has not reported yet.</span>`;
        const e = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;");
        return `<span style="font-size:12px;color:var(--muted,#567)">Excel automatic import: ${x.on === false ? "off" : "every " + e(x.every) + " min"} · last look ${e(hhmm(x.at))}${x.imported ? " · imported " + e(x.imported) : ""}${x.report ? "<br>" + e(x.report).replace(/\n/g, "<br>") : ""}</span>`; })()}
      <details><summary>Advanced</summary>${f("clientId", "App (client) ID")}${f("authority", "Sign-in authority")}</details>
      <span class="row2"><button type="button" data-x="upd">Check for updates</button>${account ? `<button type="button" data-x="out">Sign out</button>` : ""}<button type="button" data-x="cancel">Cancel</button><button type="submit">Save</button></span></form>`;
    document.body.appendChild(w);
    w.onclick = (e) => { const x = e.target.dataset && e.target.dataset.x; if (e.target === w || x === "cancel") w.remove(); if (x === "out") signOut(); if (x === "act") { if (account && !needLogin) sync("button"); else { w.remove(); signIn(); } } if (x === "upd" && window.__ssiCheckUpdate) { e.target.disabled = true; e.target.textContent = "Checking…"; window.__ssiCheckUpdate().then((r) => { e.target.textContent = r; }); } };
    w.querySelector("form").onsubmit = (e) => {
      e.preventDefault(); const o = {};
      new FormData(e.target).forEach((v, k) => { if (k === "pdfAutoOn") return; v = String(v).trim(); if (v !== String(CFG0[k] || "")) o[k] = v; });
      if (!e.target.pdfAutoOn.checked) o.pdfAuto = "off";
      const save = () => { lsSet(LSK, o); w.remove(); setMeta("dataTag", undefined).then(() => location.reload()); };
      const trim = (x) => String(x || "").replace(/^\/+|\/+$/g, "");
      // a new phone entries folder: check it is there and that Excel looks in it, or entries would wait unseen; a new PDF folder: check it is there
      const nf = trim(o.entriesFolder ?? CFG0.entriesFolder), pf = trim(o.pdfFolder ?? CFG0.pdfFolder ?? ""), vf = trim(o.invoiceFolder ?? CFG0.invoiceFolder ?? "");
      if (o.entriesFolder !== undefined) o.entriesFolder = nf;
      if (o.pdfFolder !== undefined) o.pdfFolder = pf;
      if (o.invoiceFolder !== undefined) o.invoiceFolder = vf;
      if (vf === "" && cfg().invoiceFolder) o.invoiceFolder = "";
      const chkE = nf !== trim(cfg().entriesFolder), chkP = !!pf && pf !== trim(cfg().pdfFolder), chkV = !!vf && vf !== trim(cfg().invoiceFolder);
      if (!account || (!chkE && !chkP && !chkV)) return save();
      const say = w.querySelector(".odfolder") || w.querySelector("form").appendChild(Object.assign(document.createElement("span"), { className: "odfolder" }));
      say.style.cssText = "font-size:13px;color:#b45309"; say.textContent = "Checking the folder in OneDrive…";
      (async () => {
        let msg = "";
        try { const tok = await token(); if (!tok) throw new Error("Sign in to OneDrive first.");
          if (chkE) { if (!nf || !(await g("GET", itemUrl(nf), null, "", tok))) msg = "There is no folder '" + nf + "' in your OneDrive.";
            else if (!(await g("GET", itemUrl(nf + "/excel-status.json"), null, "", tok))) msg = "Excel has not looked in '" + nf + "' yet. In Excel run SetPhoneEntriesFolder and pick the same folder, or entries will wait there unseen."; }
          if (!msg && chkP && !(await g("GET", itemUrl(pf), null, "", tok))) msg = "There is no folder '" + pf + "' in your OneDrive for the PDFs (it will be made when the first PDF is saved).";
          if (!msg && chkV && !(await g("GET", itemUrl(vf), null, "", tok))) msg = "There is no folder '" + vf + "' in your OneDrive for the invoices (it will be made when the first invoice is saved).";
        } catch (x) { msg = "Could not check the folder: " + (x.message || x); }
        if (!msg) return save();
        const key = nf + "|" + pf + "|" + vf; if (say.dataset.sure === key) return save();
        say.dataset.sure = key; say.textContent = msg + " Tap Save again to use it anyway.";
      })();
    };
  }

  // ---------- app updates ----------
  function updates() {
    if (!("serviceWorker" in navigator)) return;
    window.__ssiCheckUpdate = async () => {
      if (!navigator.onLine) return "No internet. Try again when online.";
      try {
        const reg = await navigator.serviceWorker.getRegistration(); if (!reg) return "Could not check.";
        await reg.update();
        const wk = reg.waiting || reg.installing;
        if (!wk) return "You have the latest version.";
        if (wk.state !== "installed") await new Promise((ok) => { const f = () => { if (wk.state === "installed" || wk.state === "redundant") { wk.removeEventListener("statechange", f); ok(); } }; wk.addEventListener("statechange", f); setTimeout(ok, 20000); });
        if (reg.waiting) { wk.postMessage("skip"); return "New version found. Updating…"; }
        return "You have the latest version.";
      } catch (e) { return "Could not check: " + (e.message || e); }
    };
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
  updates();
  window.__ssiSync = { sync: () => sync("manual"), state: () => ({ meta, dirty: dirty(), account: !!account, me }) };
  document.addEventListener("DOMContentLoaded", async () => {
    ui();
    await ready;
    try { await initAuth(); } catch (e) { setStatus("Sign-in could not start: " + (e.message || e), "bad"); }
    const waitApp = () => new Promise((ok) => { const t0 = Date.now(), tick = () => (window.__ssiHooks && window.__ssiHooks.ready()) || Date.now() - t0 > 15000 ? ok() : setTimeout(tick, 200); tick(); });
    await waitApp();
    sync("start");
    window.addEventListener("online", () => kick(300));
    window.addEventListener("offline", () => sync("offline"));
    // leaving the app (home button, screen off): send what is waiting at once, before the phone pauses the app
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") kick(300); else if (Object.keys(dirty()).length) { clearTimeout(kickT); sync("leave"); } });
    const every = Math.min(60, Math.max(1, Number(cfg().syncMinutes) || 2));
    setInterval(() => { if (document.visibilityState === "visible") sync("timer"); }, every * 60 * 1000);
  });
})();
