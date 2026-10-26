const $ = id => document.getElementById(id);
const send = (type, extra={}) => chrome.runtime.sendMessage({type,...extra});
let queueCache = null;
let worksheetListLoaded = false;
let worksheetListLoading = false;
function parseProfileLinks(text){
  const physical=String(text||"").split(/\r?\n/).map(x=>x.trim()).filter(x=>x&&!x.startsWith("#"));
  const records=[];const starts=(s,index)=>/^https?:\/\//i.test(s)||/^(?:[a-z0-9-]+\.)*linkedin\.com\/in\//i.test(s)||(!index&&/^\/in\//i.test(s));
  for(const part of physical){if(starts(part,records.length))records.push(part);else if(records.length)records[records.length-1]+=part;else records.push(part);}
  return records.map(x=>x.replace(/\s+/g,"")).filter(Boolean);
}
function normalizeUrl(raw){let u=String(raw||"").trim();if(/^\/in\/[^/?#]+\/?$/i.test(u))u="https://www.linkedin.com"+u;try{const x=new URL(u),m=x.pathname.match(/^\/in\/([^/]+)\/?$/i);return /(^|\.)linkedin\.com$/i.test(x.hostname)&&m?"https://www.linkedin.com/in/"+m[1]:"";}catch{return "";}}
function alignedLines(id){const raw=$(id).value;if(!raw.trim())return [];const rows=raw.split(/\r?\n/).map(x=>x.trim());while(rows.length&&rows[rows.length-1]==="")rows.pop();return rows;}
function parsedInputs(){const rawLinks=parseProfileLinks($("linksInput").value),links=rawLinks.map(normalizeUrl),emails=alignedLines("emailsInput"),jobTitles=alignedLines("jobTitlesInput"),valid=links.filter(Boolean).length;return {rawLinks,links,emails,jobTitles,valid,invalid:rawLinks.length-valid};}
function updatePreview(){
  const {rawLinks,links,emails,jobTitles,valid,invalid}=parsedInputs();const emailCount=emails.filter(Boolean).length;const jobTitleCount=jobTitles.filter(Boolean).length;
  $("linkCount").textContent=`${rawLinks.length} profile${rawLinks.length===1?"":"s"} (${valid} valid${invalid?`, ${invalid} invalid`:""})`;
  $("emailCount").textContent=emails.length?`${emailCount} email${emailCount===1?"":"s"} across ${emails.length} row${emails.length===1?"":"s"}`:"Optional";
  $("jobTitleCount").textContent=jobTitles.length?`${jobTitleCount} title${jobTitleCount===1?"":"s"} across ${jobTitles.length} row${jobTitles.length===1?"":"s"}`:"Optional";
  const box=$("countCheck");
  if(!rawLinks.length){box.textContent="Paste profile links to preview the list.";box.className="count-check";}
  else if(invalid){box.textContent=`${invalid} invalid profile link(s). Use individual LinkedIn /in/ profile URLs.`;box.className="count-check bad";}
  else if(emails.length>rawLinks.length){box.textContent=`There are ${rawLinks.length} profiles but ${emails.length} email rows. Remove the extra email rows.`;box.className="count-check bad";}
  else if(jobTitles.length>rawLinks.length){box.textContent=`There are ${rawLinks.length} profiles but ${jobTitles.length} Report’s Job Title rows. Remove the extra title rows.`;box.className="count-check bad";}
  else{box.textContent=`Ready: ${rawLinks.length} profile(s). Emails and Report’s Job Titles pair with profile rows by index; blank rows skip that field.`;box.className="count-check good";}
  const previewItems=rawLinks.slice(0,5).map((raw,i)=>`${i+1}. ${links[i]?links[i].split("/").filter(Boolean).pop():raw}  →  ${emails[i]||"(no email)"}  |  ${jobTitles[i]||"(no title check)"}`).join("\n");
  $("preview").hidden=!previewItems;$("preview").textContent=previewItems?`Preview\n${previewItems}${rawLinks.length>5?"\n…":""}`:"";
  const active=queueCache&&["running","paused"].includes(queueCache.bulkState);$("startBulk").disabled=!rawLinks.length||invalid>0||emails.length>rawLinks.length||jobTitles.length>rawLinks.length||($("worksheetSelect").value==="__new__"&&!$("newSheetName").value.trim())||active;
  return {rawLinks,links,emails,jobTitles,valid,invalid};
}
function setMessage(text,kind=""){$("message").textContent=text||"";$("message").className=`message ${kind}`.trim();}
function slug(url){try{return new URL(url).pathname.split("/").filter(Boolean).pop()||url;}catch{return url;}}
function renderQueue(q){
  queueCache=q||{};const items=Array.isArray(q?.bulkQueue)?q.bulkQueue:[],index=Number(q?.bulkIndex||0),state=q?.bulkState||"idle",opened=Math.min(index+(q?.bulkCurrent?1:0),items.length),remaining=Math.max(0,items.length-opened);
  $("queueState").textContent=({running:"Running",paused:"Paused",complete:"Complete",stopped:"Stopped",idle:"Idle"})[state]||state;
  const sheetName=q?.bulkSheetName||"";$("queueSummary").textContent=items.length?`${opened} opened · ${remaining} waiting · ${items.length} total${sheetName?` · Sheet: ${sheetName}`:""}`:"No batch loaded yet.";
  $("pauseBulk").disabled=state!=="running";$("resumeBulk").disabled=state!=="paused";$("stopBulk").disabled=!(state==="running"||state==="paused");
  const active=state==="running"||state==="paused";for(const id of ["worksheetSelect","newSheetName","outputName","delay","tabMode","autoMode"])$(id).disabled=active;
  const parsed=parsedInputs();$("startBulk").disabled=active||!parsed.rawLinks.length||parsed.invalid>0||parsed.emails.length>parsed.rawLinks.length||parsed.jobTitles.length>parsed.rawLinks.length||($("worksheetSelect").value==="__new__"&&!$("newSheetName").value.trim());
  const body=$("queueRows");body.textContent="";
  if(!items.length){const tr=document.createElement("tr"),td=document.createElement("td");td.colSpan=5;td.className="empty-row";td.textContent="Your profile queue will appear here.";tr.appendChild(td);body.appendChild(tr);return;}
  items.forEach((item,i)=>{const url=typeof item==="string"?item:item.url||"",email=typeof item==="string"?"":item.email||"",jobTitle=typeof item==="string"?"":item.report_job_title||"";const status=i<index?"Opened":(q?.bulkCurrent&&i===index?"Opening":"Waiting");const tr=document.createElement("tr");
    [String(i+1),slug(url),email,jobTitle,status].forEach((val,j)=>{const td=document.createElement("td");td.textContent=val;if(j===1)td.className="profile";if(j===4)td.className=status==="Waiting"?"status-waiting":status==="Opening"?"status-current":"status-opened";tr.appendChild(td);});body.appendChild(tr);});
}
async function refresh(){try{const [status,stored]=await Promise.all([send("LI_COPIER_STATUS"),chrome.storage.local.get({bulkQueue:[],bulkIndex:0,bulkState:"idle",bulkCurrent:null,bulkSheetName:"",delay:3,tabMode:"new",outputName:"profiles.xlsx"})]);
  if(status?.service&&!worksheetListLoaded)await loadWorksheets();
  $("serviceState").textContent=status?.service?(status.enabled?"Capture enabled":"Service online"):"Python service offline";$("serviceState").className=`service-state ${status?.service?"online":"offline"}`;renderQueue(stored);
  if(!status?.service&&!$("message").textContent)setMessage("Local Python service is offline. Start the service before opening profiles.","error");else if(status?.service?.excel_error)setMessage(`Excel update pending: ${status.service.excel_error}`,"error");
}catch(e){setMessage(`Could not refresh queue: ${e.message||e}`,"error");}}
function workbookFilename(){let name=$("outputName").value.trim()||"profiles.xlsx";if(!/\.xlsx$/i.test(name))name+=".xlsx";return name;}
async function loadWorksheets(){
  if(worksheetListLoading)return;
  worksheetListLoading=true;
  const select=$("worksheetSelect"),previous=select.value;
  const headers={"X-LI-Copier-Token":LOCAL_SERVICE_TOKEN};
  try{
    const response=await fetch(`http://127.0.0.1:8765/sheets?filename=${encodeURIComponent(workbookFilename())}`,{headers,signal:AbortSignal.timeout(8000)});
    const data=await response.json();if(!response.ok||!data.ok)throw new Error(data.error||"Could not list workbook sheets");
    worksheetListLoaded=true;
    const sheets=["Profiles",...(data.sheets||[]).filter(name=>String(name).toLowerCase()!=="profiles")];
    select.replaceChildren();
    for(const name of sheets){const option=document.createElement("option");option.value=name;option.textContent=name==="Profiles"?"Profiles (default)":name;select.appendChild(option);}
    const create=document.createElement("option");create.value="__new__";create.textContent="Create a new worksheet…";select.appendChild(create);
    select.value=[...select.options].some(option=>option.value===previous)?previous:"Profiles";
  }catch(e){setMessage(`Could not load worksheet list: ${e.message||e}. Profiles remains available.`,"error");}
  worksheetListLoading=false;
  toggleNewSheet();updatePreview();
}
function toggleNewSheet(){const create=$("worksheetSelect").value==="__new__";$("newSheetField").hidden=!create;}
chrome.storage.local.get({delay:3,tabMode:"new",autoMode:false,outputName:"profiles.xlsx"}).then(async v=>{$("delay").value=Math.max(3,Math.min(60,Number(v.delay)||3));$("tabMode").value=v.tabMode||"new";$("autoMode").checked=Boolean(v.autoMode);$("outputName").value=v.outputName||"profiles.xlsx";await loadWorksheets();});
$("autoMode").addEventListener("change",()=>chrome.storage.local.set({autoMode:$("autoMode").checked}));
$("worksheetSelect").addEventListener("change",()=>{toggleNewSheet();updatePreview();});
$("newSheetName").addEventListener("input",updatePreview);
$("outputName").addEventListener("change",loadWorksheets);
["linksInput","emailsInput","jobTitlesInput"].forEach(id=>$(id).addEventListener("input",updatePreview));
$("startBulk").addEventListener("click",async()=>{
  const c=updatePreview(),createNew=$("worksheetSelect").value==="__new__";
  let sheetName=createNew?$("newSheetName").value.trim():($("worksheetSelect").value||"Profiles");
  if(createNew&&!sheetName){setMessage("Enter a name for the new worksheet.","error");return;}
  if(!sheetName)sheetName="Profiles";
  if(!c.rawLinks.length||c.invalid||c.emails.length>c.rawLinks.length||c.jobTitles.length>c.rawLinks.length)return;
  const autoMode=$("autoMode").checked;
  const firstEmailByUrl=new Map();
  const records=c.links.map((url,i)=>{
    const email=c.emails[i]||"",key=url.toLowerCase(),normalizedEmail=email.trim().toLowerCase();
    const internal_duplicate=firstEmailByUrl.has(key)&&firstEmailByUrl.get(key)!==normalizedEmail;
    if(!firstEmailByUrl.has(key))firstEmailByUrl.set(key,normalizedEmail);
    return {url,email,report_job_title:c.jobTitles[i]||"",internal_duplicate,autoMode};
  });
  const delay=Math.max(3,Math.min(60,Number($("delay").value)||3));
  const outputName=workbookFilename();
  if(outputName.includes("/")||/[\\:*?"<>|]/.test(outputName)){setMessage("Workbook name contains an invalid character.","error");return;}
  $("startBulk").disabled=true;setMessage("Preparing the selected workbook and worksheet…");
  try{
    const headers={"Content-Type":"application/json","X-LI-Copier-Token":LOCAL_SERVICE_TOKEN};
    const configResponse=await fetch("http://127.0.0.1:8765/configure-workbook",{method:"POST",headers,body:JSON.stringify({filename:outputName})});
    const config=await configResponse.json();if(!configResponse.ok||!config.ok)throw new Error(config.error||"Could not set workbook name");
    if(createNew){
      const sheetResponse=await fetch("http://127.0.0.1:8765/create-sheet",{method:"POST",headers,body:JSON.stringify({sheet_name:sheetName})});
      const sheet=await sheetResponse.json();if(!sheetResponse.ok||!sheet.ok)throw new Error(sheet.error||"Could not create the worksheet");sheetName=sheet.sheet;
    }else{
      const listResponse=await fetch(`http://127.0.0.1:8765/sheets?filename=${encodeURIComponent(outputName)}`,{headers,signal:AbortSignal.timeout(8000)});
      const list=await listResponse.json();if(!listResponse.ok||!list.ok)throw new Error(list.error||"Could not verify worksheet");
      const existing=(list.sheets||[]).find(name=>name.toLowerCase()===sheetName.toLowerCase());
      if(sheetName!=="Profiles"&&!existing)throw new Error(`Worksheet “${sheetName}” is not in this workbook. Refresh the worksheet list and choose again.`);
      if(existing)sheetName=existing;
    }
    await chrome.storage.local.set({delay,tabMode:$("tabMode").value,autoMode:$("autoMode").checked,outputName});
    const result=await send("LI_COPIER_BULK_START",{records,delay,mode:$("tabMode").value,autoMode,sheetName});
    if(!result?.ok)throw new Error(result?.error||"Could not start the queue");
    setMessage(autoMode?`Auto Mode started for ${records.length} profile(s). Each tab will scroll to Experience, save to Excel, close, then continue.`:`Started ${records.length} profile(s) in worksheet “${sheetName}”. Tabs will open ${delay} seconds apart after each capture finishes.`,"success");
    await refresh();
  }catch(e){setMessage(`Could not start: ${e.message||e}. Check that the local service is running.`,"error");updatePreview();}
});
$("pauseBulk").addEventListener("click",async()=>{await send("LI_COPIER_BULK_PAUSE");await refresh();});$("resumeBulk").addEventListener("click",async()=>{await send("LI_COPIER_BULK_RESUME");await refresh();});$("stopBulk").addEventListener("click",async()=>{await send("LI_COPIER_BULK_STOP");await refresh();});
$("clearForm").addEventListener("click",()=>{$("linksInput").value="";$("emailsInput").value="";$("jobTitlesInput").value="";$("newSheetName").value="";setMessage("");updatePreview();});
document.addEventListener("visibilitychange",()=>{if(!document.hidden)refresh();});updatePreview();refresh();
