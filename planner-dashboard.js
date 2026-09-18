/* UB11 Schedule Planner
 * Uses the existing Google OAuth config.js when present.
 * Live source: FM tracker in Drive + Assets Inventory Ledger Google Sheet.
 */
(() => {
  "use strict";

  const PC = Object.assign({
    GOOGLE_CLIENT_ID: "",
    FM_WORKBOOK_NAMES: ["Copy ion FM tracker", "Copy of FM daily Tracker", "FM daily Tracker"],
    MAIN_SHEET_NAME: "FM collection tracker",
    INVENTORY_LEDGER_SHEET_ID: "1VC44seK6vR2IHl53Bn0NwentqRd8XwhSZJ0RtWB67uU",
    INVENTORY_LEDGER_NAME: "Assets Inventory Ledger",
    INVENTORY_SHEET_NAME: "Asset Inventory",
    TRANSACTIONS_SHEET_NAME: "Asset Transactions",
    CLIENT_SNAPSHOT_SHEET_NAME: "Client Inventory Snapshot",
    DAILY_SNAPSHOT_SHEET_NAME: "Daily Snapshots",
    PLANNER_SHEET_NAME: "Planner Plans",
    DRIVE_FOLDER_ID: "15NJveQRQ0WHR5FuR0E967dGs5bOBaPoH",
    OAUTH_SCOPES: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive"
  }, window.CONFIG || {});

  const CFG = {
    clientId: PC.GOOGLE_CLIENT_ID,
    fmNames: PC.FM_WORKBOOK_NAMES || [PC.PRIMARY_WORKBOOK_NAME, PC.FALLBACK_WORKBOOK_NAME].filter(Boolean),
    fmSheet: PC.MAIN_SHEET_NAME,
    ledgerId: PC.INVENTORY_LEDGER_SHEET_ID,
    ledgerName: PC.INVENTORY_LEDGER_NAME,
    inventorySheet: PC.INVENTORY_SHEET_NAME,
    transactionSheet: PC.TRANSACTIONS_SHEET_NAME,
    clientSnapshotSheet: PC.CLIENT_SNAPSHOT_SHEET_NAME,
    dailySnapshotSheet: PC.DAILY_SNAPSHOT_SHEET_NAME,
    plannerSheet: PC.PLANNER_SHEET_NAME,
    folderId: PC.DRIVE_FOLDER_ID || PC.DRIVE_PHOTOS_PARENT_FOLDER_ID,
    scope: PC.OAUTH_SCOPES
  };

  const state = {
    view: "day",
    date: ymd(new Date()),
    shift: "all",
    sort: "effective",
    filter: "all",
    records: [],
    allRecords: [],
    inventory: [],
    transactions: [],
    plans: [],
    connected: false,
    fmFile: null,
    token: null,
    tokenExpiresAt: 0,
    tokenClient: null,
    pendingAuth: null,
    initialized: false
  };

  const el = id => document.getElementById(id);
  const root = el("plannerApp");
  const modalBack = el("modalBack");
  const modal = el("modal");
  const toastEl = el("toast");

  const FM_HEADERS = [
    "HSC","LOAD TYPE","Code","Cleint ","Broker","Post Code","Vehicle Type","Planned arrival ","Actual Arrival ",
    "Planned Pull","Actual Pull","ETA to HSC","Actual arrival time HSC","Pallets","Bags / loose parcles ","Status","Reg","Driver Name","In APP"
  ];

  const snapshotAssets = ["Black Pallets","Wooden Sleeves","Wood Pallets","THG Pallets","Magnum Lids","Red Sleeves","Magnum","Car Bags","Magnum Base","Purple Bags"];

  function log(...args){ console.debug("[UB11 Planner]", ...args); }
  function toast(message, good=true){ toastEl.textContent = message; toastEl.classList.toggle("show", true); toastEl.style.background = good ? "#16202a" : "#7f1d1d"; clearTimeout(toast._t); toast._t = setTimeout(()=>toastEl.classList.remove("show"), 4200); }
  function escapeHtml(s){ return String(s ?? "").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;","\"":"&quot;"}[c])); }
  function pad(n){ return String(n).padStart(2,"0"); }
  function ymd(d){ return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }
  function fromYmd(s){ const [y,m,d] = String(s).split("-").map(Number); return new Date(y,m-1,d); }
  function isoWeek(d){ const x = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const day = x.getUTCDay() || 7; x.setUTCDate(x.getUTCDate() + 4 - day); const yearStart = new Date(Date.UTC(x.getUTCFullYear(),0,1)); return Math.ceil((((x-yearStart)/86400000)+1)/7); }
  function fmtDate(s, opts={weekday:"short", day:"2-digit", month:"short", year:"numeric"}){ return fromYmd(s).toLocaleDateString("en-GB", opts); }
  function londonYmd(v){
    const d = v instanceof Date ? v : new Date(v);
    return new Intl.DateTimeFormat("en-CA", {timeZone:"Europe/London",year:"numeric",month:"2-digit",day:"2-digit"}).format(d);
  }
  function hm(d){ return new Intl.DateTimeFormat("en-GB", {hour:"2-digit", minute:"2-digit", hour12:false, timeZone:"Europe/London"}).format(d); }
  function excelSerialToDate(serial, baseDate){
    const n = Number(serial);
    if (!Number.isFinite(n)) return null;
    const dt = fromYmd(baseDate);
    const frac = ((n % 1) + 1) % 1;
    dt.setHours(0,0,0,0);
    dt.setMinutes(Math.round(frac * 1440));
    return dt;
  }
  function excelSerialToYmd(n){
    const d = new Date(Date.UTC(1899,11,30) + Number(n) * 86400000);
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`;
  }
  function ymdToExcelSerial(s){ return Math.round((fromYmd(s) - new Date(1899,11,30)) / 86400000); }
  function clean(v){ return String(v ?? "").trim(); }
  function num(v){ const n=Number(v); return Number.isFinite(n) ? n : 0; }
  function hasValue(v){ return v !== undefined && v !== null && String(v).trim() !== ""; }
  function htmlTable(headers, rows){
    return `<div class="table-wrap"><table class="table"><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>${rows.length ? rows.map(r=>`<tr>${r.map(c=>`<td>${escapeHtml(c)}</td>`).join("")}</tr>`).join("") : `<tr><td colspan="${headers.length}">No data</td></tr>`}</tbody></table></div>`;
  }

  function setStatus(text, good=false){
    el("syncStatus").textContent = text;
    el("syncStatus").style.color = good ? "#0f8a6a" : "#6b7785";
  }

  function authAvailable(){ return Boolean(CFG.clientId && window.google?.accounts?.oauth2); }

  function loadGoogleIdentityServices(){
    if(window.google?.accounts?.oauth2) return Promise.resolve();
    if(window.__ub11GisPromise) return window.__ub11GisPromise;
    window.__ub11GisPromise = new Promise((resolve,reject)=>{
      const existing=document.querySelector('script[data-ub11-gis]');
      if(existing){
        const started=Date.now();
        const poll=()=>{
          if(window.google?.accounts?.oauth2) return resolve();
          if(Date.now()-started>10000) return reject(new Error("Google Identity Services did not finish loading. Check that https://accounts.google.com is not blocked by the browser/network."));
          setTimeout(poll,100);
        };
        poll();
        return;
      }
      const script=document.createElement('script');
      script.src='https://accounts.google.com/gsi/client';
      script.async=true; script.defer=true; script.dataset.ub11Gis='1';
      script.onload=()=>{
        const started=Date.now();
        const poll=()=>{
          if(window.google?.accounts?.oauth2) return resolve();
          if(Date.now()-started>10000) return reject(new Error("Google Identity Services loaded, but the OAuth API is unavailable."));
          setTimeout(poll,50);
        };
        poll();
      };
      script.onerror=()=>reject(new Error("Could not load Google Identity Services. Check your network, content blockers, or firewall."));
      document.head.appendChild(script);
    });
    return window.__ub11GisPromise;
  }

  async function getToken(){
    if(state.token && Date.now() < state.tokenExpiresAt - 45000) return state.token;
    if(!CFG.clientId) throw new Error("config.js was not loaded. Put config.js in the same folder as planner-dashboard.html.");
    if(location.protocol === 'file:') throw new Error("Do not open planner-dashboard.html with a file:// URL. Serve it from the same http/https origin as your existing FM dashboard so Google OAuth can authorise it.");
    await loadGoogleIdentityServices();
    if(!authAvailable()) throw new Error("Google OAuth is unavailable after loading config.js and Google Identity Services. Check the Google OAuth client ID and authorised JavaScript origins.");
    return new Promise((resolve,reject)=>{
      state.pendingAuth = {resolve,reject};
      try{
        state.tokenClient ||= google.accounts.oauth2.initTokenClient({
          client_id: CFG.clientId,
          scope: CFG.scope,
          callback: (response)=>{
            if(response.error){ state.pendingAuth?.reject(new Error(response.error_description || response.error)); state.pendingAuth=null; return; }
            state.token = response.access_token;
            state.tokenExpiresAt = Date.now() + (Number(response.expires_in || 3600) * 1000);
            state.connected = true;
            setStatus("Google Drive connected", true);
            state.pendingAuth?.resolve(state.token); state.pendingAuth=null;
          }
        });
        state.tokenClient.requestAccessToken({prompt:""});
      }catch(err){ state.pendingAuth?.reject(err); state.pendingAuth=null; }
    });
  }

  async function api(url, options={}){
    let token = await getToken();
    const run = async t => fetch(url, Object.assign({}, options, {headers:Object.assign({Authorization:`Bearer ${t}`}, options.headers||{})}));
    let res = await run(token);
    if(res.status===401){ state.token=null; token=await getToken(); res=await run(token); }
    if(!res.ok){ const txt=await res.text(); throw new Error(`${res.status}: ${txt.slice(0,500)}`); }
    const type=res.headers.get("content-type")||"";
    return type.includes("application/json") ? res.json() : res.arrayBuffer();
  }

  function driveQueryName(name){
    return `'${String(name).replace(/'/g,"\\'")}'`;
  }
  async function driveListByName(name){
    const q = `name = ${driveQueryName(name)} and trashed = false`;
    const u = `https://www.googleapis.com/drive/v3/files?${new URLSearchParams({q,fields:"files(id,name,mimeType,modifiedTime,webViewLink,parents)",pageSize:"20",supportsAllDrives:"true",includeItemsFromAllDrives:"true"})}`;
    return (await api(u)).files || [];
  }

  async function driveCopy(fileId, name){
    const u = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}/copy?${new URLSearchParams({supportsAllDrives:"true",fields:"id,name,mimeType,webViewLink,parents"})}`;
    return api(u,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name,parents:[CFG.folderId]})});
  }

  async function sheetsGet(spreadsheetId, range, valueRenderOption="UNFORMATTED_VALUE"){
    const u = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?${new URLSearchParams({valueRenderOption})}`;
    return (await api(u)).values || [];
  }

  async function sheetsPut(spreadsheetId, range, values){
    const u = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?${new URLSearchParams({valueInputOption:"RAW"})}`;
    return api(u,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify({range,majorDimension:"ROWS",values})});
  }

  async function sheetsAppend(spreadsheetId, range, values){
    const u = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}:append?${new URLSearchParams({valueInputOption:"RAW",insertDataOption:"INSERT_ROWS"})}`;
    return api(u,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({range,majorDimension:"ROWS",values})});
  }

  async function sheetsClear(spreadsheetId, range){
    const u = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}:clear`;
    return api(u,{method:"POST",headers:{"Content-Type":"application/json"},body:"{}"});
  }

  async function sheetsMetadata(spreadsheetId){
    return api(`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?${new URLSearchParams({fields:"sheets.properties"})}`);
  }

  async function sheetsBatchUpdate(spreadsheetId, requests){
    return api(`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}:batchUpdate`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({requests})});
  }

  async function ensureSheet(title, headers){
    const meta = await sheetsMetadata(CFG.ledgerId);
    const found = (meta.sheets||[]).find(s=>s.properties?.title===title);
    if(!found){
      await sheetsBatchUpdate(CFG.ledgerId,[{addSheet:{properties:{title}}}]);
      await sheetsPut(CFG.ledgerId, `${title}!A1:${columnLetter(headers.length)}1`, [headers]);
    } else if(headers?.length){
      const existing = await sheetsGet(CFG.ledgerId, `${title}!A1:${columnLetter(headers.length)}2`);
      const flat = existing.flat().map(clean);
      const missing = headers.some((h,i)=>clean(flat[i])!==h);
      if(missing){ await sheetsPut(CFG.ledgerId, `${title}!A1:${columnLetter(headers.length)}1`, [headers]); }
    }
  }
  function columnLetter(n){ let s=""; while(n>0){ const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); } return s; }

  async function readFMWorkbook(){
    let files=[];
    for(const name of CFG.fmNames){
      files = await driveListByName(name);
      if(files.length) break;
    }
    if(!files.length) throw new Error(`Could not find the FM tracker in Drive. Tried: ${CFG.fmNames.join(", ")}`);
    const preferred = files.find(f=>f.mimeType==="application/vnd.google-apps.spreadsheet") || files[0];
    state.fmFile = preferred;
    setStatus(`FM source: ${preferred.name}`);
    if(preferred.mimeType==="application/vnd.google-apps.spreadsheet"){
      return sheetsGet(preferred.id, `${CFG.fmSheet}!A1:S2000`);
    }
    // Supports uploaded XLSX/Excel files in Drive too.
    const buf = await api(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(preferred.id)}?alt=media`);
    const XLSX = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");
    const wb = XLSX.read(buf,{type:"array",cellDates:false});
    const ws = wb.Sheets[CFG.fmSheet] || wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""}).slice(0,2000);
  }

  function isDateHeaderRow(row){
    if(!row) return null;
    const n=Number(row[0]);
    const day=clean(row[1]).toUpperCase();
    if(Number.isFinite(n) && /^(MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)$/.test(day)) return excelSerialToYmd(n);
    return null;
  }

  function parseFM(rows){
    const out=[]; let currentDate=null; let currentHsc=""; let currentType="";
    const headers = rows?.[0] || FM_HEADERS;
    for(let i=1;i<rows.length;i++){
      const row=rows[i]||[];
      const headerDate=isDateHeaderRow(row);
      if(headerDate){ currentDate=headerDate; currentHsc=""; currentType=""; continue; }
      if(!currentDate) continue;
      if(hasValue(row[0])) currentHsc=clean(row[0]);
      if(hasValue(row[1])) currentType=clean(row[1]);
      const meaningful = [row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[11],row[12],row[13],row[14],row[16],row[17]].some(hasValue);
      if(!meaningful || !currentHsc) continue;
      if(currentHsc.toUpperCase() !== "UB11") continue;
      const routeType=currentType.toUpperCase();
      const client=clean(row[3]);
      const routeRaw=clean(row[5]);
      let category="inbound";
      if(routeType.includes("INTERCITY")) category="intercity";
      else if(/RETURN|RTS|OUTBOUND/.test(routeType)) category="outbound";
      const routeParts=routeRaw.split(/\s*->\s*/).map(clean).filter(Boolean);
      let route="";
      if(routeParts.length>=2){ route=`${routeParts[0]} → ${routeParts[1]}`; }
      else if(category==="outbound") route=`UB11 → ${client || routeRaw || "destination"}`;
      else route=`${client || routeRaw || "Origin"} → UB11`;
      const planned = excelSerialToDate(row[7],currentDate);
      const actual = excelSerialToDate(row[12],currentDate) || excelSerialToDate(row[8],currentDate);
      const effective = actual || planned;
      const d={
        row:i+1,date:currentDate,hsc:"UB11",type:routeType,code:clean(row[2]),client,broker:clean(row[4]),postcode:routeRaw,vehicle:clean(row[6]),
        plannedArrival:planned,actualArrival:actual,plannedPull:excelSerialToDate(row[9],currentDate),actualPull:excelSerialToDate(row[10],currentDate),eta:excelSerialToDate(row[11],currentDate),
        pallets:clean(row[13]),bags:clean(row[14]),status:clean(row[15]),reg:clean(row[16]),driver:clean(row[17]),inApp:clean(row[18]),category,route,effectiveTime:effective
      };
      // Ignore empty continuation rows with no time/client/code after inherited HSC.
      if(!d.client && !d.code && !d.reg && !d.plannedArrival && !d.postcode) continue;
      out.push(d);
    }
    return out;
  }

  function effectiveShiftLabel(date, time){
    if(!time) return "";
    const ds=ymd(time), t=time.getHours()+time.getMinutes()/60;
    if(ds===date && t>=7 && t<19) return "day";
    const next=ymd(new Date(fromYmd(date).getTime()+86400000));
    if((ds===date && t>=19)||(ds===next && t<7)) return "night";
    return "other";
  }

  function selectedRecords(){
    const selected=state.date;
    let rows=[];
    if(state.shift==="night"){
      const next=ymd(new Date(fromYmd(selected).getTime()+86400000));
      rows=state.allRecords.filter(r=>r.date===selected||r.date===next).filter(r=>effectiveShiftLabel(selected,r.effectiveTime)==="night");
    } else if(state.shift==="day"){
      rows=state.allRecords.filter(r=>r.date===selected).filter(r=>effectiveShiftLabel(selected,r.effectiveTime)==="day");
    } else {
      rows=state.allRecords.filter(r=>r.date===selected);
    }
    if(state.filter!=="all") rows=rows.filter(r=>r.category===state.filter);
    const sorters={
      effective:(a,b)=>(a.effectiveTime?.getTime()||0)-(b.effectiveTime?.getTime()||0),
      client:(a,b)=>a.client.localeCompare(b.client),type:(a,b)=>a.type.localeCompare(b.type),vehicle:(a,b)=>a.vehicle.localeCompare(b.vehicle),status:(a,b)=>a.status.localeCompare(b.status)
    };
    rows.sort(sorters[state.sort]||sorters.effective);
    return rows;
  }

  function renderKpis(records){
    const count=(c)=>records.filter(r=>r.category===c).length;
    el("kpis").innerHTML=[
      ["Total movements",records.length,"UB11 only"],["Inbound",count("inbound"),"Client → UB11"],["Outbound",count("outbound"),"UB11 → destination"],["Intercity",count("intercity"),"Intercity route"]
    ].map(x=>`<div class="kpi"><div class="label">${x[0]}</div><div class="value">${x[1]}</div><div class="sub">${x[2]}</div></div>`).join("");
  }

  function renderDay(){
    const rec=selectedRecords(); state.records=rec; renderKpis(rec);
    if(!rec.length){ el("main").innerHTML=`<div class="empty"><strong>No UB11 movements found</strong><div class="hint">${escapeHtml(fmtDate(state.date))}. Try All shifts or refresh the Drive data.</div></div>`; return; }
    const top=state.shift==="night"?"19:00–07:00":state.shift==="day"?"07:00–19:00":"Today";
    const hours=state.shift==="night"?Array.from({length:12},(_,i)=>(19+i)%24):Array.from({length:12},(_,i)=>7+i);
    const hourNodes=hours.map(h=>`<div class="hour">${pad(h)}:00</div>`).join("");
    const cards=rec.map((r,idx)=>{
      const cls=r.category; const planned=r.plannedArrival?hm(r.plannedArrival):"—"; const actual=r.actualArrival?hm(r.actualArrival):"—"; const movement=r.effectiveTime?hm(r.effectiveTime):"—";
      const load=[r.pallets && `${r.pallets} pallets`,r.bags && r.bags].filter(Boolean).join(" • ") || "—";
      const status=r.status || (r.inApp?"In APP":"Planned");
      return `<article class="delivery ${cls}">
        <div class="d-top"><span class="badge ${cls}">${escapeHtml(cls.toUpperCase())}</span><span class="badge">${escapeHtml(r.type || "MOVEMENT")}</span><span class="route">${escapeHtml(r.route)}</span><span class="time">${escapeHtml(movement)}</span></div>
        <div class="d-grid">
          <div class="kv"><small>Client</small><strong>${escapeHtml(r.client||"—")}</strong></div><div class="kv"><small>FM Ref</small><strong>${escapeHtml(r.code||"—")}</strong></div>
          <div class="kv"><small>Planned arrival</small><strong>${escapeHtml(planned)}</strong></div><div class="kv"><small>Actual arrival</small><strong>${escapeHtml(actual)}</strong></div>
          <div class="kv"><small>Postcode / route</small><strong>${escapeHtml(r.postcode||"—")}</strong></div><div class="kv"><small>Vehicle</small><strong>${escapeHtml(r.vehicle||"—")}</strong></div>
          <div class="kv"><small>Reg</small><strong>${escapeHtml(r.reg||"—")}</strong></div><div class="kv"><small>Driver</small><strong>${escapeHtml(r.driver||"—")}</strong></div>
          <div class="kv"><small>Load</small><strong>${escapeHtml(load)}</strong></div><div class="kv"><small>Planned pull</small><strong>${escapeHtml(r.plannedPull?hm(r.plannedPull):"—")}</strong></div>
          <div class="kv"><small>Actual pull</small><strong>${escapeHtml(r.actualPull?hm(r.actualPull):"—")}</strong></div><div class="kv"><small>Status</small><strong>${escapeHtml(status)}</strong></div>
        </div>
      </article>`;
    }).join("");
    el("main").innerHTML=`<div class="day-shell"><div class="timeline">${hourNodes}</div><div class="slots"><div style="padding:12px 12px 0;font-size:12px;font-weight:750;color:#566271">${escapeHtml(fmtDate(state.date,{weekday:"long",day:"numeric",month:"long",year:"numeric"}))} • ${escapeHtml(top)}</div><div class="cards">${cards}</div></div></div>`;
  }

  function startOfWeek(s){ const d=fromYmd(s); const day=d.getDay()||7; const x=new Date(d); x.setDate(d.getDate()-day+1); return ymd(x); }
  function renderWeek(){
    const start=startOfWeek(state.date); const days=Array.from({length:7},(_,i)=>{const d=new Date(fromYmd(start)); d.setDate(d.getDate()+i); return ymd(d)});
    renderKpis(state.allRecords.filter(r=>days.includes(r.date)));
    const cells=days.map((d,i)=>{
      const rec=state.allRecords.filter(r=>r.date===d); const plans=state.plans.filter(p=>p.date===d); const sunday=i===6;
      return `<div class="day-cell ${d===state.date?"" : ""}">
        <div class="day-num">${escapeHtml(fmtDate(d,{weekday:"short",day:"2-digit",month:"short"}))}</div>
        <div class="day-count">${rec.length} movement${rec.length===1?"":"s"}</div>
        <div class="pills">${rec.slice(0,4).map(r=>`<span class="pill">${escapeHtml(r.effectiveTime?hm(r.effectiveTime):"—")} ${escapeHtml(r.category)}</span>`).join("")}${rec.length>4?`<span class="pill">+${rec.length-4} more</span>`:""}${plans.map(p=>`<span class="pill plan">Plan: ${escapeHtml(p.title||p.type)}</span>`).join("")}</div>
        <div class="cell-actions"><button class="btn small" data-open-day="${d}">Open</button><button class="btn small" data-plan-date="${d}">Plan</button>${sunday?`<button class="btn small primary" data-snapshot-week="${d}">Week snapshot</button>`:""}</div>
      </div>`;
    }).join("");
    el("main").innerHTML=`<div class="calendar"><div class="cal-title"><h2>Week ${isoWeek(fromYmd(state.date))}, ${fromYmd(start).getFullYear()}</h2><span class="hint">Sunday = week-end inventory snapshot</span></div><div class="week-grid">${cells}</div></div>`;
  }

  function renderMonth(){
    const d=fromYmd(state.date); const year=d.getFullYear(), month=d.getMonth(); const first=new Date(year,month,1); const firstDay=(first.getDay()||7)-1; const daysIn=new Date(year,month+1,0).getDate();
    renderKpis(state.allRecords.filter(r=>r.date.startsWith(`${year}-${pad(month+1)}`)));
    const cells=[]; for(let i=0;i<firstDay;i++) cells.push(`<div class="day-cell outside"></div>`);
    for(let day=1;day<=daysIn;day++){
      const s=`${year}-${pad(month+1)}-${pad(day)}`; const rec=state.allRecords.filter(r=>r.date===s); const plans=state.plans.filter(p=>p.date===s); const audit=plans.find(p=>String(p.type).toUpperCase().includes("AUDIT"));
      cells.push(`<div class="day-cell"><div class="day-num">${day}</div><div class="day-count">${rec.length} movement${rec.length===1?"":"s"}</div><div class="pills">${audit?`<span class="pill audit">Inventory audit</span>`:""}${plans.filter(p=>p!==audit).map(p=>`<span class="pill plan">${escapeHtml(p.title||p.type)}</span>`).join("")}</div><div class="cell-actions"><button class="btn small" data-open-day="${s}">Open</button><button class="btn small" data-plan-date="${s}" data-plan-type="INVENTORY AUDIT">Audit plan</button></div></div>`);
    }
    while(cells.length%7) cells.push(`<div class="day-cell outside"></div>`);
    const headings=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].map(x=>`<div class="weekday">${x}</div>`).join("");
    el("main").innerHTML=`<div class="calendar"><div class="cal-title"><h2>${escapeHtml(d.toLocaleString("en-GB",{month:"long",year:"numeric"}))}</h2><span class="hint">Select a day to plan an inventory audit</span></div><div class="week-grid">${headings}${cells.join("")}</div></div>`;
  }

  function renderYear(){
    const y=fromYmd(state.date).getFullYear(); renderKpis(state.allRecords.filter(r=>r.date.startsWith(`${y}-`)));
    const cards=Array.from({length:12},(_,m)=>{const prefix=`${y}-${pad(m+1)}`; const rec=state.allRecords.filter(r=>r.date.startsWith(prefix)); const audits=state.plans.filter(p=>p.date?.startsWith(prefix)&&String(p.type).toUpperCase().includes("AUDIT")); const inbound=rec.filter(r=>r.category==="inbound").length; const outbound=rec.filter(r=>r.category==="outbound").length; const intercity=rec.filter(r=>r.category==="intercity").length; return `<div class="month-card"><h3>${new Date(y,m,1).toLocaleString("en-GB",{month:"long"})}</h3><div class="mbar"><span>Total</span><b>${rec.length}</b></div><div class="mbar"><span>Inbound</span><b>${inbound}</b></div><div class="mbar"><span>Outbound</span><b>${outbound}</b></div><div class="mbar"><span>Intercity</span><b>${intercity}</b></div><div class="mbar"><span>Audit plans</span><b>${audits.length}</b></div><div style="margin-top:8px"><button class="btn small" data-open-month="${y}-${pad(m+1)}-01">Open month</button></div></div>`;}).join("");
    el("main").innerHTML=`<div class="calendar" style="padding:14px"><div class="cal-title" style="padding:0 0 14px;border-bottom:0"><h2>${y}</h2><span class="hint">Planner overview</span></div><div class="year-grid">${cards}</div></div>`;
  }

  function render(){
    document.querySelectorAll("#viewSeg button").forEach(b=>b.classList.toggle("active",b.dataset.view===state.view));
    el("datePicker").value=state.date; el("shiftSelect").value=state.shift; el("sortSelect").value=state.sort; el("typeFilter").value=state.filter;
    if(state.view==="day") renderDay(); else if(state.view==="week") renderWeek(); else if(state.view==="month") renderMonth(); else renderYear();
  }

  async function loadData(){
    setStatus("Refreshing Drive data…");
    try{
      const fmRows=await readFMWorkbook(); state.allRecords=parseFM(fmRows);
      const inventoryRows=await sheetsGet(CFG.ledgerId, `${CFG.inventorySheet}!A1:B500`); state.inventory=inventoryRows.slice(1).filter(r=>clean(r[0]));
      const transactionRows=await sheetsGet(CFG.ledgerId, `${CFG.transactionSheet}!A1:H10000`); state.transactions=transactionRows.slice(1).filter(r=>clean(r[0]));
      try{ await ensureSheet(CFG.plannerSheet,["ID","Date","Type","Shift","Title","Notes","Created","User"]); }catch(e){ log("Planner sheet could not be prepared",e); }
      try{ state.plans=parsePlans(await sheetsGet(CFG.ledgerId, `${CFG.plannerSheet}!A1:H5000`)); }catch(e){ state.plans=[]; }
      state.connected=true; setStatus(`Live • ${state.allRecords.length} UB11 movements loaded`,true); render();
    }catch(err){
      console.error(err); setStatus("Data connection needed");
      el("main").innerHTML=`<div class="empty"><strong>Connect Google Drive to load the live schedule</strong><div class="hint">${escapeHtml(err.message)}</div><div style="margin-top:12px"><button class="btn primary" id="emptyConnect">Connect Google Drive</button></div></div>`;
      el("emptyConnect")?.addEventListener("click",connectAndLoad);
    }
  }

  function parsePlans(rows){
    return (rows||[]).slice(1).filter(r=>clean(r[1])).map((r,i)=>({id:clean(r[0])||`row-${i}`,date:clean(r[1]),type:clean(r[2]),shift:clean(r[3]),title:clean(r[4]),notes:clean(r[5]),created:clean(r[6]),user:clean(r[7])}));
  }

  function openModal(title,body,foot=""){
    modal.innerHTML=`<div class="modal-head"><h3>${escapeHtml(title)}</h3><button class="btn small" data-close-modal>Close</button></div><div class="modal-body">${body}</div>${foot?`<div class="modal-foot">${foot}</div>`:""}`; modalBack.classList.add("open");
  }
  function closeModal(){modalBack.classList.remove("open"); modal.innerHTML="";}

  async function savePlan(){
    const date=el("planDate")?.value; const type=el("planType")?.value; const shift=el("planShift")?.value; const title=clean(el("planTitle")?.value); const notes=clean(el("planNotes")?.value);
    if(!date||!title){ toast("Date and plan title are required.",false); return; }
    const id=`PLAN-${Date.now()}`;
    const now=new Date().toISOString();
    try{ await sheetsAppend(CFG.ledgerId, `${CFG.plannerSheet}!A:H`, [[id,date,type,shift,title,notes,now,""]]); state.plans.push({id,date,type,shift,title,notes,created:now,user:""}); toast("Plan saved."); closeModal(); render(); }
    catch(e){ toast(`Could not save plan: ${e.message}`,false); }
  }

  function openPlanModal(date,type="OPERATIONS"){
    const defaults=type||"OPERATIONS";
    openModal("Add plan",`<div class="form-grid">
      <div class="form-row"><label>Date</label><input id="planDate" type="date" value="${escapeHtml(date)}"></div>
      <div class="form-row"><label>Plan type</label><select id="planType"><option ${defaults==="OPERATIONS"?"selected":""}>OPERATIONS</option><option ${defaults==="INVENTORY AUDIT"?"selected":""}>INVENTORY AUDIT</option><option>OTHER</option></select></div>
      <div class="form-row"><label>Shift</label><select id="planShift"><option>ALL</option><option>DAY</option><option>NIGHT</option></select></div>
      <div class="form-row"><label>Title</label><input id="planTitle" placeholder="e.g. Inventory audit"></div>
      <div class="form-row full"><label>Notes</label><textarea id="planNotes" rows="4" placeholder="Plan details"></textarea></div>
      </div>`,`<button class="btn" data-close-modal>Cancel</button><button class="btn primary" id="savePlanBtn">Save plan</button>`);
    el("savePlanBtn").addEventListener("click",savePlan);
  }

  async function runWeekSnapshot(sunday){
    const d=fromYmd(sunday); const week=isoWeek(d); const name=`${week}_${d.getFullYear()}`;
    try{
      const existing=(await api(`https://www.googleapis.com/drive/v3/files?${new URLSearchParams({q:`'${CFG.folderId}' in parents and name = '${name}' and trashed = false`,fields:"files(id,name,webViewLink,mimeType)",pageSize:"10",supportsAllDrives:"true",includeItemsFromAllDrives:"true"})}`)).files||[];
      if(existing.length){ toast(`Week snapshot ${name} already exists.`); openModal("Week snapshot already exists",`<p>${escapeHtml(name)} is already in the snapshot folder.</p><p><a href="${escapeHtml(existing[0].webViewLink||`https://drive.google.com/drive/folders/${CFG.folderId}`)}" target="_blank" rel="noreferrer">Open snapshot</a></p>`); return; }
      const copy=await driveCopy(CFG.ledgerId,name); toast(`Week snapshot ${name} created.`); openModal("Week snapshot created",`<p>Created <strong>${escapeHtml(name)}</strong> in the Drive snapshot folder.</p>${copy.webViewLink?`<p><a href="${escapeHtml(copy.webViewLink)}" target="_blank" rel="noreferrer">Open copied workbook</a></p>`:""}`);
    }catch(e){ toast(`Week snapshot failed: ${e.message}`,false); }
  }

  async function loadInventoryForRoutine(){
    await getToken();
    const [inventoryRows, transactionRows] = await Promise.all([
      sheetsGet(CFG.ledgerId, `${CFG.inventorySheet}!A1:B500`),
      sheetsGet(CFG.ledgerId, `${CFG.transactionSheet}!A1:H10000`)
    ]);
    const inventory = inventoryRows.slice(1).filter(r=>clean(r[0])).map(r=>({asset:clean(r[0]),balance:num(r[1])}));
    const transactions = transactionRows.slice(1).filter(r=>clean(r[0]));
    if(!inventory.length) throw new Error(`No rows were returned from "${CFG.inventorySheet}". Check the sheet name and Google access.`);
    if(!transactionRows.length) throw new Error(`No rows were returned from "${CFG.transactionSheet}". Check the sheet name and Google access.`);
    return {inventory,transactions};
  }

  function aggregateDayTransactions(transactions,date){
    const agg={};
    transactions.filter(r=>londonYmd(r[0])===date).forEach(r=>{
      const client=clean(r[1]);
      const movement=clean(r[2]).toUpperCase();
      const asset=clean(r[3]);
      const q=num(r[4]);
      if(!client||!asset||!['SENT','RECEIVED'].includes(movement)) return;
      agg[client] ||= {};
      agg[client][asset] ||= {sent:0,received:0};
      if(movement==='SENT') agg[client][asset].sent += q;
      if(movement==='RECEIVED') agg[client][asset].received += q;
    });
    return agg;
  }

  async function runEndShiftRoutine(){
    try{
      setStatus("Loading inventory snapshot…");
      const {inventory,transactions}=await loadInventoryForRoutine();
      state.inventory=inventory.map(x=>[x.asset,x.balance]);
      state.transactions=transactions;
      const nowLondonHour=Number(new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/London',hour:'2-digit',hour12:false}).format(new Date()));
      const shift=state.shift==='all' ? (nowLondonHour>=7&&nowLondonHour<19?'Day':'Night') : (state.shift==='day'?'Day':'Night');
      const agg=aggregateDayTransactions(transactions,state.date);
      const clients=Object.keys(agg).sort();
      const clientRows=clients.map(c=>{
        let totalS=0,totalR=0;
        Object.values(agg[c]).forEach(x=>{ totalS+=x.sent; totalR+=x.received; });
        return [c,totalS,totalR,totalS-totalR];
      });
      const txDateCount=transactions.filter(r=>londonYmd(r[0])===state.date).length;
      openModal("End shift routine",`
        <div class="hint">Date <strong>${escapeHtml(state.date)}</strong> • Shift <strong>${escapeHtml(shift)}</strong>. This writes the current Asset Inventory balances to Daily Snapshots and replaces this day's client snapshot with the transactions recorded for this date only.</div>
        <h4>Warehouse snapshot — ${escapeHtml(shift)}</h4>${htmlTable(["Asset","Current balance"],inventory.map(x=>[x.asset,x.balance]))}
        <h4>Client movement snapshot — ${escapeHtml(state.date)}</h4><div class="hint">${txDateCount} transaction${txDateCount===1?'':'s'} found for this date.</div>${htmlTable(["Client","Sent","Received","Sent − Received"],clientRows)}
        `,`<button class="btn" data-close-modal>Cancel</button><button class="btn primary" id="confirmEodBtn">Update snapshots</button>`);
      el("confirmEodBtn").addEventListener("click",()=>commitEndShift(shift,inventory,agg));
      setStatus("Ready to update snapshots",true);
    }catch(e){
      console.error("End shift routine",e);
      toast(`End shift routine could not load: ${e.message}`,false);
      setStatus("End shift routine failed");
    }
  }

  async function commitEndShift(shift,inventory,agg){
    const btn=el("confirmEodBtn"); if(btn) btn.disabled=true;
    try{
      await getToken();
      const dailyHeaderRows=await sheetsGet(CFG.ledgerId,`${CFG.dailySnapshotSheet}!A1:L4`);
      const headerRow=(dailyHeaderRows.find(r=>clean(r[0]).toUpperCase()==='SHIFT')||[]);
      let assetHeaders=headerRow.slice(2).filter(Boolean);
      if(!assetHeaders.length){
        assetHeaders=inventory.map(x=>x.asset);
        await sheetsPut(CFG.ledgerId,`${CFG.dailySnapshotSheet}!A2:${columnLetter(assetHeaders.length+2)}2`,[["Shift","Date",...assetHeaders]]);
      }
      const byAsset=Object.fromEntries(inventory.map(x=>[x.asset,x.balance]));
      const row=[shift,ymdToExcelSerial(state.date),...assetHeaders.map(a=>byAsset[a] ?? "")];
      const existing=await sheetsGet(CFG.ledgerId,`${CFG.dailySnapshotSheet}!A3:${columnLetter(row.length)}5000`);
      const dateSerial=ymdToExcelSerial(state.date);
      let target=-1;
      existing.forEach((r,i)=>{ if(clean(r[0]).toUpperCase()===shift.toUpperCase() && Number(r[1])===dateSerial) target=i+3; });
      if(target>0) await sheetsPut(CFG.ledgerId,`${CFG.dailySnapshotSheet}!A${target}:${columnLetter(row.length)}${target}`,[row]);
      else {
        let last=2;
        existing.forEach((r,i)=>{if(r.some(hasValue)) last=i+3;});
        const dest=last+1;
        await sheetsPut(CFG.ledgerId,`${CFG.dailySnapshotSheet}!A${dest}:${columnLetter(row.length)}${dest}`,[row]);
      }

      await ensureSheet(CFG.clientSnapshotSheet,["Date","Client",...assetHeaders]);
      const clientWidth=assetHeaders.length+2;
      const existingClient=await sheetsGet(CFG.ledgerId,`${CFG.clientSnapshotSheet}!A3:${columnLetter(clientWidth)}5000`);
      const current=Object.keys(agg).sort();
      const oldRows=[];
      existingClient.forEach((r,i)=>{if(Number(r[0])===dateSerial) oldRows.push(i+3);});
      const snapRows=current.map(client=>[dateSerial,client,...assetHeaders.map(asset=>{const x=agg[client]?.[asset]; return x ? x.sent-x.received : 0;})]);
      const work=Math.max(oldRows.length,snapRows.length);
      let appendRow=2;
      existingClient.forEach((r,j)=>{if(r.some(hasValue)) appendRow=j+3;});
      appendRow += 1;
      for(let i=0;i<work;i++){
        const rowNo=oldRows[i] || null;
        if(rowNo && i<snapRows.length) await sheetsPut(CFG.ledgerId,`${CFG.clientSnapshotSheet}!A${rowNo}:${columnLetter(clientWidth)}${rowNo}`,[snapRows[i]]);
        else if(rowNo) await sheetsClear(CFG.ledgerId,`${CFG.clientSnapshotSheet}!A${rowNo}:${columnLetter(clientWidth)}${rowNo}`);
        else if(i<snapRows.length) {
          const dest=appendRow++;
          await sheetsPut(CFG.ledgerId,`${CFG.clientSnapshotSheet}!A${dest}:${columnLetter(clientWidth)}${dest}`,[snapRows[i]]);
        }
      }
      toast(`End shift complete — ${shift} snapshot saved for ${state.date}.`,true);
      closeModal();
      setStatus("Snapshots updated",true);
    }catch(e){
      console.error("Snapshot update",e);
      toast(`Snapshot update failed: ${e.message}`,false);
      if(btn) btn.disabled=false;
    }
  }

  async function openEodReport(){
    try{
      if(!state.transactions.length||!state.inventory.length) await loadData();
      const tx=state.transactions.filter(r=>londonYmd(r[0])===state.date); const rec=tx.filter(r=>clean(r[2]).toUpperCase()==="RECEIVED"); const sent=tx.filter(r=>clean(r[2]).toUpperCase()==="SENT");
      const netAssets={}; tx.forEach(r=>{const a=clean(r[3]); if(!a) return; netAssets[a] ||= {received:0,sent:0}; const k=clean(r[2]).toLowerCase(); if(netAssets[a][k]!==undefined) netAssets[a][k]+=num(r[4]);});
      const rows=Object.entries(netAssets).map(([a,x])=>[a,x.received,x.sent,x.received-x.sent]);
      openModal(`EOD report — ${state.date}`,`<div class="report"><div class="report-summary"><div class="report-card"><small>Transactions</small><b>${tx.length}</b></div><div class="report-card"><small>Received</small><b>${rec.length}</b></div><div class="report-card"><small>Sent</small><b>${sent.length}</b></div></div><h4>Asset movement</h4>${htmlTable(["Asset","Received","Sent","Net to warehouse"],rows)}<h4>Current warehouse inventory</h4>${htmlTable(["Asset","Balance"],state.inventory.map(r=>[clean(r[0]),num(r[1])]))}</div>`,`<button class="btn" data-close-modal>Close</button><button class="btn primary" id="printReportBtn">Print report</button>`);
      el("printReportBtn").addEventListener("click",()=>window.print());
    }catch(e){ toast(`EOD report failed: ${e.message}`,false); }
  }

  async function connectAndLoad(){
    try{ setStatus("Connecting to Google…"); await getToken(); await loadData(); }
    catch(e){ setStatus("Connection failed"); toast(`Google connection failed: ${e.message}`,false); }
  }

  function shiftDateStep(direction){
    const d=fromYmd(state.date); if(state.view==="day") d.setDate(d.getDate()+direction); else if(state.view==="week") d.setDate(d.getDate()+7*direction); else if(state.view==="month") d.setMonth(d.getMonth()+direction); else d.setFullYear(d.getFullYear()+direction); state.date=ymd(d); render();
  }

  root.addEventListener("click", async e=>{
    const t=e.target.closest("button,#endShiftLink,#eodLink,[data-open-day],[data-plan-date],[data-snapshot-week],[data-open-month],[data-close-modal]"); if(!t) return;
    if(t.dataset.view){ state.view=t.dataset.view; render(); return; }
    if(t.id==="connectBtn"){ connectAndLoad(); return; }
    if(t.id==="refreshBtn"){ loadData(); return; }
    if(t.id==="prevBtn"){ shiftDateStep(-1); return; }
    if(t.id==="nextBtn"){ shiftDateStep(1); return; }
    if(t.id==="todayBtn"){ state.date=ymd(new Date()); render(); return; }
    if(t.id==="endShiftLink"){ runEndShiftRoutine(); return; }
    if(t.id==="eodLink"){ openEodReport(); return; }
    if(t.dataset.openDay){ state.date=t.dataset.openDay; state.view="day"; render(); return; }
    if(t.dataset.planDate){ openPlanModal(t.dataset.planDate,t.dataset.planType||"OPERATIONS"); return; }
    if(t.dataset.snapshotWeek){ await runWeekSnapshot(t.dataset.snapshotWeek); return; }
    if(t.dataset.openMonth){ state.date=t.dataset.openMonth; state.view="month"; render(); return; }
    if(t.dataset.closeModal){ closeModal(); return; }
  });

  modalBack.addEventListener("click",e=>{ if(e.target===modalBack || e.target.closest("[data-close-modal]")) closeModal(); });
  el("datePicker").addEventListener("change",e=>{state.date=e.target.value;render();});
  el("shiftSelect").addEventListener("change",e=>{state.shift=e.target.value;render();});
  el("sortSelect").addEventListener("change",e=>{state.sort=e.target.value;render();});
  el("typeFilter").addEventListener("change",e=>{state.filter=e.target.value;render();});

  function init(){
    if(state.initialized) return; state.initialized=true; el("datePicker").value=state.date; render();
    if(!CFG.clientId) setStatus("config.js not loaded");
    else if(location.protocol === 'file:') setStatus("Serve from your allowed web origin");
    else setStatus("Ready to connect to Google");
  }
  init();
})();
