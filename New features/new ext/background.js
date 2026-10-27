importScripts("config.js");
const SERVICE = "http://127.0.0.1:8765";
const PROFILE_URLS = ["https://*.linkedin.com/in/*"];
const PROFILE_RE = /^https:\/\/(?:[a-z]{1,3}\.)?linkedin\.com\/in\//i;
const REQUEST_TIMEOUT_MS = 8000;
const MAX_CONCURRENT_CAPTURES = 1;
let bulkQueue = [];
let bulkState = "idle";
let bulkIndex = 0;
let bulkDelay = 3000;
let bulkMode = "new";
let bulkAutoMode = false;
let bulkMapping = {};
let bulkSheetName = "";
let bulkCurrent = null;
let bulkHasOpened = false;
let bulkTabRecordsCache = {};
let bulkTimer = null;
const reported404Tabs = new Set();
let keywordsCache = null;   // read once; kept fresh by the onChanged listener below
async function getKeywords() {
  if (keywordsCache === null) keywordsCache = (await chrome.storage.local.get({keywords:[]})).keywords || [];
  return keywordsCache;
}
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === "local" && changes.keywords) keywordsCache = changes.keywords.newValue || [];
});
const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
async function normalizeProfileUrl(raw) {
  let value = String(raw || "").trim();
  if (/^\/in\/[^/?#]+\/?$/i.test(value)) value = "https://www.linkedin.com" + value;
  try {
    const u = new URL(value);
    if (!/(^|\.)linkedin\.com$/i.test(u.hostname)) return "";
    const m = u.pathname.match(/^\/in\/([^/]+)\/?$/i);
    return m ? "https://www.linkedin.com/in/" + m[1] : "";
  } catch (_) { return ""; }
}
let bulkStateLoaded = false;
async function ensureBulkState() {
  if (bulkStateLoaded) return;
  const saved = await chrome.storage.local.get({bulkQueue:[],bulkIndex:0,bulkState:"idle",bulkDelay:3000,bulkMode:"new",bulkAutoMode:false,bulkMapping:{},bulkSheetName:"",bulkCurrent:null});
  bulkQueue=saved.bulkQueue || []; bulkIndex=Number(saved.bulkIndex || 0); bulkState=saved.bulkState || "idle";
  bulkDelay=Number(saved.bulkDelay || 3000); bulkMode=saved.bulkMode || "new"; bulkAutoMode=Boolean(saved.bulkAutoMode); bulkMapping=saved.bulkMapping || {};
  bulkSheetName=saved.bulkSheetName || ""; bulkCurrent=saved.bulkCurrent || null; bulkHasOpened=saved.bulkHasOpened == null ? bulkIndex>0 : Boolean(saved.bulkHasOpened);
  if (bulkCurrent) { bulkIndex=Math.max(bulkIndex,Number(bulkCurrent.index||0)+1); bulkCurrent=null; bulkHasOpened=true; await chrome.storage.local.set({bulkIndex,bulkCurrent:null,bulkHasOpened}); }
  bulkStateLoaded=true;
}
async function queueStatus() {
  await ensureBulkState();
  if (!bulkQueue.length) {
    const saved = await chrome.storage.local.get({bulkQueue:[],bulkIndex:0,bulkState:"idle"});
    if (saved.bulkQueue?.length) { bulkQueue=saved.bulkQueue; bulkIndex=Number(saved.bulkIndex||0); bulkState=saved.bulkState||"idle"; }
  }
  return {state:bulkState,total:bulkQueue.length,index:bulkIndex,remaining:Math.max(0,bulkQueue.length-bulkIndex),current:bulkCurrent};
}
function scheduleQueueAdvance(delayMs) {
  if (bulkTimer) clearTimeout(bulkTimer);
  bulkTimer = null;
  chrome.alarms.clear("li-copier-next");
  const delay = Math.max(3000, Math.min(60000, Number(delayMs) || 3000));
  if (delay < 30000) {
    bulkTimer = setTimeout(() => {
      bulkTimer = null;
      runBulkQueue();
    }, delay);
  } else {
    chrome.alarms.create("li-copier-next", {delayInMinutes:delay / 60000});
  }
}

async function saveBulkTabRecord(tabId, record) {
  if (tabId == null) return;
  const saved = await chrome.storage.session.get({bulkTabRecords:{}});
  const records = saved.bulkTabRecords || {};
  records[tabId] = record;
  await chrome.storage.session.set({bulkTabRecords:records});
}
async function bulkTabRecord(tabId) {
  if (tabId == null) return null;
  const saved = await chrome.storage.session.get({bulkTabRecords:{}});
  return saved.bulkTabRecords?.[tabId] || null;
}

async function runBulkQueue() {
  if (bulkState !== "running") return;
  if (bulkIndex >= bulkQueue.length) {
    bulkState = "complete";
    bulkCurrent = null;
    await chrome.storage.local.set({bulkState,bulkIndex,bulkQueue,bulkCurrent:null});
    chrome.alarms.clear("li-copier-next");
    chrome.alarms.clear("li-copier-capture-timeout");
    return;
  }
  const rec = bulkQueue[bulkIndex];
  const url = await normalizeProfileUrl(rec.url);
  if (!url) {
    recordQueueFailure({...rec,sheet_name:rec.sheet_name || bulkSheetName}, "Invalid Link", "Input is not a valid LinkedIn /in/ profile URL.");
    bulkIndex++;
    await chrome.storage.local.set({bulkIndex,bulkQueue,bulkState,bulkDelay,bulkMode,bulkMapping});
    scheduleQueueAdvance(bulkDelay);
    return;
  }
  const tabRecord = {profileKey:profileKey(url), url, email:rec.email || "", report_job_title:rec.report_job_title || "", target_company:rec.target_company || "", internal_duplicate:Boolean(rec.internal_duplicate), autoMode:Boolean(rec.autoMode || bulkAutoMode), sheet_name:rec.sheet_name || bulkSheetName || "Profiles"};
  bulkMapping[url.toLowerCase()] = {email:rec.email || "", target_company:rec.target_company || "", originalUrl:rec.url, sheet_name:tabRecord.sheet_name};
  try {
    let tab = null;
    if (!tabRecord.autoMode && bulkMode === "reuse") {
      const tabs = await chrome.tabs.query({url:["https://*.linkedin.com/in/*"]});
      tab = tabs.find(t => (t.url || "").split("?")[0].replace(/\/$/,"").toLowerCase() === url.toLowerCase() && (!bulkHasOpened || !t.active));
    }
    let tabId;
    if (tabRecord.autoMode) {
      // Register the per-tab batch data before navigating so the content script
      // knows to scroll to Experience and close only this batch-created tab.
      const opened = await chrome.tabs.create({url:"about:blank",active:true});
      tabId = opened.id;
      await saveBulkTabRecord(tabId, tabRecord);
      bulkTabRecordsCache[tabId] = tabRecord;
      await chrome.tabs.update(tabId,{url,active:true});
      check404Redirect(tabId, url);
    } else if (tab?.id != null) {
      tabId = tab.id;
      await saveBulkTabRecord(tabId, tabRecord);
      await chrome.tabs.update(tabId,{url,active:!bulkHasOpened});
      check404Redirect(tabId, url);
    } else {
      const opened = await chrome.tabs.create({url,active:!bulkHasOpened});
      tabId = opened.id;
      await saveBulkTabRecord(tabId, tabRecord);
      check404Redirect(tabId, url);
    }
    bulkHasOpened = true;
    bulkTabRecordsCache[tabId] = tabRecord;
    bulkMapping[url.toLowerCase()] = {...bulkMapping[url.toLowerCase()],tabId};
    bulkIndex++;
    bulkCurrent = null;
    await chrome.storage.local.set({bulkCurrent:null,bulkQueue,bulkIndex,bulkState,bulkDelay,bulkMode,bulkMapping,bulkSheetName,bulkHasOpened});
  } catch (e) {
    recordQueueFailure({...rec,url,sheet_name:tabRecord.sheet_name}, "Capture Failed", String(e));
    bulkIndex++;
    bulkCurrent = null;
    await chrome.storage.local.set({bulkCurrent:null,bulkIndex,bulkQueue,bulkState,bulkDelay,bulkMode,bulkMapping,bulkSheetName,bulkHasOpened});
  }
  // Opening proceeds on schedule even if the tab later redirects to LinkedIn's 404 page
  // or its capture is still running in the background.
  if (!tabRecord.autoMode) scheduleQueueAdvance(bulkDelay);
}

async function saveAutoTabAndContinue(tabId, record) {
  if (!record?.autoMode || tabId == null) return false;
  try {
    const response = await fetch(`${SERVICE}/export`, {
      method:"POST", headers:{"X-LI-Copier-Token":LOCAL_SERVICE_TOKEN},
      signal:AbortSignal.timeout(30000)
    });
    let result = {};
    try { result = await response.json(); } catch (_) {}
    if (!response.ok || !result.ok) throw new Error(result.error || `Excel save returned ${response.status}`);
    try { await chrome.tabs.remove(tabId); } catch (_) {}
    if (bulkState === "running") scheduleQueueAdvance(bulkDelay);
    return true;
  } catch (error) {
    try { await chrome.tabs.sendMessage(tabId,{type:"LI_COPIER_AUTO_SAVE_FAILED",reason:String(error)}); } catch (_) {}
    console.warn("[Experience Copier] Auto Mode kept the tab open because Excel did not confirm the save:",String(error));
    return false;
  }
}

async function recordQueueFailure(rec,status,comment) {
  try {
    const response = await fetch(`${SERVICE}/record-status`,{method:"POST",headers:{"Content-Type":"application/json","X-LI-Copier-Token":LOCAL_SERVICE_TOKEN},
      body:JSON.stringify({url:rec.url,email:rec.email,report_job_title:rec.report_job_title,target_company:rec.target_company,capture_status:status,comment,sheet_name:rec.sheet_name,
        internal_duplicate:Boolean(rec.internal_duplicate)}),signal:AbortSignal.timeout(5000)});
    return response.ok;
  } catch (_) { return false; }
}


// ---- Remember WHEN each profile tab was opened, so Excel rows can follow the order the links were opened.
// Stored in chrome.storage.session (survives the service worker sleeping; cleared when the browser closes).
const OPEN_KEY = "openedTabs";
let stampChain = Promise.resolve();
const queueStamp = fn => { stampChain = stampChain.then(fn).catch(() => {}); return stampChain; };

function profileKey(u) {
  try {
    const parts = new URL(u).pathname.split("/").filter(Boolean);
    return parts[0] === "in" && parts[1] ? "/in/" + parts[1].toLowerCase() : "";
  } catch (_) { return ""; }
}
async function getOpened() {
  return (await chrome.storage.session.get({ [OPEN_KEY]: {} }))[OPEN_KEY];
}
async function stampTabNow(tabId, url) {
  const key = profileKey(url);
  if (!key || tabId == null) return;
  const map = await getOpened();
  if (map[tabId]?.key === key) return;                       // same profile in the same tab: keep first time
  const newest = Math.max(0, ...Object.values(map).map(v => v.stamp));
  map[tabId] = { key, stamp: Math.max(Date.now(), newest + 1) };   // strictly increasing
  await chrome.storage.session.set({ [OPEN_KEY]: map });
}
chrome.tabs.onCreated.addListener(tab => {
  const url = tab.pendingUrl || tab.url || "";
  if (profileKey(url)) queueStamp(() => stampTabNow(tab.id, url));
});
chrome.tabs.onUpdated.addListener((tabId, info) => {
  if (info.url && profileKey(info.url)) queueStamp(() => stampTabNow(tabId, info.url));
});
async function openedAtFor(tabId, url) {
  await stampChain;                                           // let any just-created tab finish stamping
  const entry = (await getOpened())[tabId];
  return entry && entry.key === profileKey(url) ? entry.stamp : null;
}

// A slow profile request must not hold every other profile in a serial queue.
// The service stages captures under a lock and performs Excel writes separately.
let activeCaptureUploads = 0;
const captureUploadQueue = [];
function startCaptureUploads() {
  while (activeCaptureUploads < MAX_CONCURRENT_CAPTURES && captureUploadQueue.length) {
    const job = captureUploadQueue.shift();
    activeCaptureUploads++;
    const started = Date.now();
    fetch(`${SERVICE}/capture`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-LI-Copier-Token": LOCAL_SERVICE_TOKEN },
      body: JSON.stringify(job.payload || {}),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS)
    }).then(async r => {
      let result = {};
      try { result = await r.json(); } catch (_) {}
      return { ok: r.ok, statusCode: r.status, result };
    }).then(value => {
      console.info(`[Experience Copier] Excel staging response in ${Date.now() - started} ms`);
      job.resolve(value);
    }, error => {
      console.warn(`[Experience Copier] Excel staging request failed after ${Date.now() - started} ms`, String(error));
      job.reject(error);
    }).finally(() => {
      activeCaptureUploads--;
      startCaptureUploads();
    });
  }
}
function enqueueCapture(payload) {
  return new Promise((resolve, reject) => {
    captureUploadQueue.push({ payload, resolve, reject });
    startCaptureUploads();
  });
}

async function injectAndInit(tabId, enabled) {
  try {
    await chrome.scripting.executeScript({ target: { tabId }, files: ["content.js"] });
    await chrome.tabs.sendMessage(tabId, { type: "LI_COPIER_SET_ENABLED", enabled });
    return true;
  } catch (_) { return false; }
}

async function injectProfileTabs(enabled) {
  const tabs = await chrome.tabs.query({ url: PROFILE_URLS });
  const leftToRight = [...tabs].sort((a, b) => a.windowId - b.windowId || a.index - b.index);
  queueStamp(async () => { for (const t of leftToRight) await stampTabNow(t.id, t.url); });
  await Promise.all(tabs.filter(t => t.id != null).map(t => injectAndInit(t.id, enabled)));
  return tabs.length;
}

async function handleCaptureTerminal(tabId, message) {
  if (tabId == null || !message.url || !profileKey(message.url) || message.captureOk) return;
  const record = await bulkTabRecord(tabId) || {url:message.url,email:"",sheet_name:"Profiles"};
  const recorded = await recordQueueFailure(record, "Unable to Capture", message.reason || "Capture could not be completed.");
  if (record.autoMode && recorded) await saveAutoTabAndContinue(tabId,record);
}

chrome.alarms.onAlarm.addListener(async alarm => {
  await ensureBulkState();
  if (alarm.name !== "li-copier-next") return;
  const saved = await chrome.storage.local.get({bulkQueue:[],bulkIndex:0,bulkState:"idle",bulkDelay:3000,bulkMode:"new",bulkAutoMode:false,bulkMapping:{},bulkSheetName:"",bulkCurrent:null});
  bulkQueue = saved.bulkQueue || []; bulkIndex = Number(saved.bulkIndex || 0); bulkState = saved.bulkState || "idle";
  bulkDelay = Number(saved.bulkDelay || 3000); bulkMode = saved.bulkMode || "new"; bulkAutoMode=Boolean(saved.bulkAutoMode); bulkMapping = saved.bulkMapping || {}; bulkSheetName = saved.bulkSheetName || ""; bulkCurrent=saved.bulkCurrent || null;
  if (bulkState === "running" && !bulkCurrent) await runBulkQueue();
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  (async () => {
    if (message?.type === "LI_COPIER_SLEEP") {          // un-throttled timer for background tabs
      await wait(Math.min(Math.max(Number(message.ms) || 0, 0), 5000));
      sendResponse({ ok: true }); return;
    }
    await ensureBulkState();
    if (message?.type === "LI_COPIER_START") {
      await chrome.storage.local.set({ captureEnabled: true });
      const tabsFound = await injectProfileTabs(true);
      const [active] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
      if (active?.id && PROFILE_RE.test(active.url || "")) {
        await injectAndInit(active.id, true);
        try { await chrome.tabs.sendMessage(active.id, { type: "LI_COPIER_CAPTURE_NOW" }); } catch (_) {}
      }
      sendResponse({ ok: true, tabsFound }); return;
    }
    if (message?.type === "LI_COPIER_STOP") {
      await chrome.storage.local.set({ captureEnabled: false });
      const tabsFound = await injectProfileTabs(false);
      sendResponse({ ok: true, tabsFound }); return;
    }
    if (message?.type === "LI_COPIER_GET_TAB_OPTIONS") {
      const record = await bulkTabRecord(sender?.tab?.id);
      const currentKey = profileKey(sender?.tab?.url || "");
      sendResponse({autoMode:Boolean(record?.autoMode && currentKey && currentKey === record.profileKey)}); return;
    }
    if (message?.type === "LI_COPIER_CAPTURE_RESULT") {
      if (message.terminal) await handleCaptureTerminal(sender?.tab?.id, message);
      sendResponse({ ok: true });
      return;
    }
    if (message?.type === "LI_COPIER_SUBMIT_CAPTURE") {
      try {
        const payload = message.payload || {};
        const tabId = sender?.tab?.id;
        const [keywords, item, opened] = await Promise.all([
          getKeywords(), bulkTabRecord(tabId), openedAtFor(tabId, payload.url || "")
        ]);
        const captureKey = profileKey(payload.url || "");
        if (item && item.profileKey === captureKey) {
          payload.sheet_name = item.sheet_name;
          payload.email = item.email || "";
          payload.report_job_title = item.report_job_title || "";
          payload.target_company = item.target_company || "";
          payload.internal_duplicate = Boolean(item.internal_duplicate);
        } else {
          payload.email = "";
        }
        payload.keywords = keywords;
        payload.opened_at = opened ?? payload.first_seen ?? Date.now();
        // Return as soon as the service journals the capture. Excel flush confirmation
        // runs in a separate message so this profile tab and the batch can continue.
        const staged = await enqueueCapture(payload);
        sendResponse(staged);
        if (staged?.ok && staged.result?.ok && item?.autoMode && String(payload.experience_text || "").trim()) {
          await saveAutoTabAndContinue(tabId,item);
        }
      } catch (e) {
        sendResponse({ ok: false, statusCode: 0, error: "Python service unavailable: " + String(e) });
      }
      return;
    }
    if (message?.type === "LI_COPIER_BULK_START") {
      if (bulkState === "running" || bulkState === "paused") {
        sendResponse({ok:false,error:"A batch is already active. Stop it before starting another."}); return;
      }
      bulkSheetName = String(message.sheetName || "Profiles").trim() || "Profiles";
      bulkQueue = (Array.isArray(message.records) ? message.records : []).map(rec => ({...rec,sheet_name:bulkSheetName}));
      bulkIndex = 0;
      bulkDelay = Math.max(3000, Math.min(60000, Number(message.delay || 3) * 1000));
      bulkAutoMode = Boolean(message.autoMode || (Array.isArray(message.records) && message.records.some(rec => rec?.autoMode)));
      bulkMode = message.mode === "reuse" ? "reuse" : "new";
      bulkState = "running";
      bulkTabRecordsCache = {}; bulkCurrent = null; bulkHasOpened = false; bulkStateLoaded = true;
      chrome.alarms.clear("li-copier-capture-timeout");
      await chrome.storage.local.set({bulkIndex:0,bulkState,bulkCurrent:null,bulkHasOpened:false,bulkTotal:bulkQueue.length,bulkSheetName,captureEnabled:true});
      if (bulkTimer) clearTimeout(bulkTimer);
      await chrome.storage.local.set({bulkQueue,bulkIndex,bulkState,bulkDelay,bulkMode,bulkAutoMode,bulkMapping,bulkSheetName,bulkHasOpened});
      runBulkQueue();
      sendResponse({ok:true,queue:await queueStatus(),sheetName:bulkSheetName}); return;
    }
    if (message?.type === "LI_COPIER_BULK_PAUSE") {
      bulkState = "paused"; if (bulkTimer) clearTimeout(bulkTimer); chrome.alarms.clear("li-copier-next");
      await chrome.storage.local.set({bulkState});
      sendResponse({ok:true,queue:await queueStatus()}); return;
    }
    if (message?.type === "LI_COPIER_BULK_RESUME") {
      if (bulkIndex < bulkQueue.length) { bulkState = "running"; await chrome.storage.local.set({bulkState}); runBulkQueue(); }
      sendResponse({ok:true,queue:await queueStatus()}); return;
    }
    if (message?.type === "LI_COPIER_BULK_STOP") {
      bulkState = "stopped"; if (bulkTimer) clearTimeout(bulkTimer); chrome.alarms.clear("li-copier-next"); chrome.alarms.clear("li-copier-capture-timeout");
      await chrome.storage.local.set({bulkState});
      sendResponse({ok:true,queue:await queueStatus()}); return;
    }
    if (message?.type === "LI_COPIER_BULK_STATUS") {
      sendResponse({queue:await queueStatus()}); return;
    }
    if (message?.type === "LI_COPIER_STATUS") {
      const enabled = Boolean((await chrome.storage.local.get({ captureEnabled: false })).captureEnabled);
      let service = null;
      try {
        const r = await fetch(`${SERVICE}/status`, {
          headers: { "X-LI-Copier-Token": LOCAL_SERVICE_TOKEN },
          signal: AbortSignal.timeout(5000)
        });
        if (r.ok) service = await r.json();
      } catch (_) {}
      sendResponse({ enabled, service }); return;
    }
    if (message?.type === "LI_COPIER_EXPORT") {
      try {
        const r = await fetch(`${SERVICE}/export`, {
          method: "POST",
          headers: { "X-LI-Copier-Token": LOCAL_SERVICE_TOKEN },
          signal: AbortSignal.timeout(30000)  // a forced Excel write can take a moment
        });
        sendResponse({ ok: r.ok, result: await r.json() });
      } catch (e) { sendResponse({ ok: false, error: String(e) }); }
    }
  })();
  return true;
});

async function check404Redirect(tabId, requestedUrl="") {
  let currentUrl=requestedUrl;
  try { const tab=await chrome.tabs.get(tabId); currentUrl=tab.url || currentUrl; } catch (_) {}
  if (!/linkedin\.com\/404\/?(?:[?#]|$)/i.test(currentUrl)) return;
  const record=await bulkTabRecord(tabId);
  const key=`${tabId}:${record?.profileKey || ""}`;
  if (record && !reported404Tabs.has(key)) {
    reported404Tabs.add(key);
    const recorded = await recordQueueFailure(record,"Invalid Link","LinkedIn redirected this URL to /404/.");
    if (record.autoMode && recorded) await saveAutoTabAndContinue(tabId,record);
  }
}
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  const target = changeInfo.url || tab.url || "";
  if (/linkedin\.com\/404\/?(?:[?#]|$)/i.test(target)) check404Redirect(tabId,target);

});
chrome.tabs.onRemoved.addListener(tabId => {
  delete bulkTabRecordsCache[tabId];
  chrome.storage.session.get({bulkTabRecords:{}}).then(saved => {
    const records = saved.bulkTabRecords || {};
    if (records[tabId]) { delete records[tabId]; return chrome.storage.session.set({bulkTabRecords:records}); }
  }).catch(() => {});
});
