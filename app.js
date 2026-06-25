(function(){
  "use strict";

  var ADDIN_NAMESPACE = "fleetdocsfines";
  var pageState = null;
  var apiRef = null;
  var assets = [];
  var documents = [];
  var currentFile = null;
  var currentOcrText = "";
  var ghIndexSha = null;

  function byId(id){ return document.getElementById(id); }
  function esc(v){ return String(v == null ? "" : v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;").replace(/'/g,"&#039;"); }
  function norm(v){ return String(v || "").trim().toUpperCase().replace(/[^A-Z0-9]/g,""); }
  function todayIso(){ return new Date().toISOString().slice(0,10); }
  function nowStamp(){ return new Date().toISOString().replace(/[-:]/g,"").replace(/\..+/,""); }
  function uid(prefix){ return (prefix || "doc") + "_" + nowStamp() + "_" + Math.random().toString(36).slice(2,8); }
  function getCfg(){ return { owner:byId("ghOwner").value.trim(), repo:byId("ghRepo").value.trim(), branch:byId("ghBranch").value.trim() || "main", token:byId("ghToken").value.trim() }; }
  function setStatus(txt){ byId("ocrStatus").textContent = txt; }

  function saveLocalCfg(){
    var cfg = getCfg();
    localStorage.setItem("fleetDocsGhOwner", cfg.owner);
    localStorage.setItem("fleetDocsGhRepo", cfg.repo);
    localStorage.setItem("fleetDocsGhBranch", cfg.branch);
    setStatus("Configurazione locale salvata. Il token non viene salvato.");
  }
  function loadLocalCfg(){
    byId("ghOwner").value = localStorage.getItem("fleetDocsGhOwner") || "";
    byId("ghRepo").value = localStorage.getItem("fleetDocsGhRepo") || "";
    byId("ghBranch").value = localStorage.getItem("fleetDocsGhBranch") || "main";
  }

  function apiGet(typeName, search, limit){
    return new Promise(function(resolve,reject){
      if(!apiRef){ resolve([]); return; }
      apiRef.call("Get", { typeName:typeName, search:search || {}, resultsLimit:limit || 5000 }, function(r){ resolve(r || []); }, reject);
    });
  }
  function entityId(e){ if(!e)return""; if(typeof e === "string") return e; if(e.id){ if(typeof e.id === "string") return e.id; if(e.id.id) return e.id.id; return String(e.id); } return ""; }
  function plateOf(d){ return d.licensePlate || d.licencePlate || ""; }
  function nameOf(d){ return d.name || d.serialNumber || d.vehicleIdentificationNumber || entityId(d); }
  function isActiveDevice(d){
    if(!d || entityId(d)==="NoDeviceId") return false;
    var now = new Date(), from = d.activeFrom ? new Date(d.activeFrom) : null, to = d.activeTo ? new Date(d.activeTo) : null;
    if(from && from > now) return false; if(to && to < now) return false; return true;
  }
  async function loadAssets(){
    try{ assets = (await apiGet("Device", {}, 5000)).filter(isActiveDevice); }
    catch(e){ console.warn(e); assets = []; }
    renderAssetSelect();
  }
  function renderAssetSelect(selectedId){
    var sel = byId("assetSelect");
    sel.innerHTML = '<option value="">Nessun asset selezionato</option>' + assets.map(function(a){
      var id = entityId(a); return '<option value="'+esc(id)+'" '+(id===selectedId?'selected':'')+'>'+esc(nameOf(a))+' — '+esc(plateOf(a) || "senza targa")+'</option>';
    }).join("");
  }
  function matchAssetByPlate(plate){
    var p = norm(plate); if(!p) return null;
    return assets.find(function(a){ return norm(plateOf(a)) === p; }) || null;
  }

  function ghHeaders(){
    var cfg = getCfg();
    if(!cfg.owner || !cfg.repo || !cfg.token) throw new Error("Configura owner, repository e token GitHub.");
    return { "Accept":"application/vnd.github+json", "Authorization":"Bearer "+cfg.token, "X-GitHub-Api-Version":"2022-11-28" };
  }
  function ghUrl(path){ var c=getCfg(); return "https://api.github.com/repos/"+encodeURIComponent(c.owner)+"/"+encodeURIComponent(c.repo)+"/contents/"+path.replace(/^\//,"")+"?ref="+encodeURIComponent(c.branch || "main"); }
  async function ghGetFile(path){
    var res = await fetch(ghUrl(path), { headers: ghHeaders() });
    if(res.status === 404) return { exists:false, sha:null, content:null };
    if(!res.ok) throw new Error("GitHub GET fallita: "+res.status+" "+await res.text());
    var json = await res.json();
    var content = decodeURIComponent(escape(atob((json.content || "").replace(/\n/g,""))));
    return { exists:true, sha:json.sha, content:content };
  }
  async function ghPutFile(path, content, message, sha){
    var c=getCfg();
    var body = { message: message, content: btoa(unescape(encodeURIComponent(content))), branch: c.branch || "main" };
    if(sha) body.sha = sha;
    var putUrl = "https://api.github.com/repos/"+encodeURIComponent(c.owner)+"/"+encodeURIComponent(c.repo)+"/contents/"+path.replace(/^\//,"");
    var res = await fetch(putUrl, { method:"PUT", headers: Object.assign({"Content-Type":"application/json"}, ghHeaders()), body: JSON.stringify(body) });
    if(!res.ok) throw new Error("GitHub PUT fallita: "+res.status+" "+await res.text());
    return await res.json();
  }
  async function ghPutBinary(path, file, message){
    var c=getCfg();
    var arr = await file.arrayBuffer();
    var bytes = new Uint8Array(arr); var binary = "";
    for(var i=0;i<bytes.length;i++) binary += String.fromCharCode(bytes[i]);
    var body = { message: message, content: btoa(binary), branch: c.branch || "main" };
    var putUrl = "https://api.github.com/repos/"+encodeURIComponent(c.owner)+"/"+encodeURIComponent(c.repo)+"/contents/"+path.replace(/^\//,"");
    var res = await fetch(putUrl, { method:"PUT", headers: Object.assign({"Content-Type":"application/json"}, ghHeaders()), body: JSON.stringify(body) });
    if(!res.ok) throw new Error("GitHub upload file fallito: "+res.status+" "+await res.text());
    return await res.json();
  }
  function rawUrl(path){ var c=getCfg(); return "https://raw.githubusercontent.com/"+encodeURIComponent(c.owner)+"/"+encodeURIComponent(c.repo)+"/"+encodeURIComponent(c.branch || "main")+"/"+path.replace(/^\//,""); }

  async function loadArchive(){
    setStatus("Carico archivio da GitHub...");
    await loadAssets();
    var f = await ghGetFile("data/documents.json");
    ghIndexSha = f.sha;
    documents = f.exists && f.content ? JSON.parse(f.content) : [];
    renderAll();
    setStatus("Archivio caricato: "+documents.length+" documenti.");
  }

  function typeLabel(t){ return ({fine:"Multa",registration:"Libretto",insurance:"Assicurazione",leasing:"Leasing",inspection:"Revisione",fuel_card:"Carta carburante",other:"Altro"})[t] || t || "Documento"; }
  function folderForType(t){ return ({fine:"multe",registration:"libretti",insurance:"assicurazioni",leasing:"leasing",inspection:"revisioni",fuel_card:"carte-carburante",other:"altro"})[t] || "altro"; }
  function selectedAsset(){ var id = byId("assetSelect").value; return assets.find(function(a){ return entityId(a) === id; }) || null; }
  function extensionOf(file){ var n=file && file.name || "document.pdf"; var m=n.match(/\.([a-zA-Z0-9]+)$/); return (m ? m[1] : "pdf").toLowerCase(); }

  async function runOcr(){
    currentFile = byId("fileInput").files[0];
    if(!currentFile){ alert("Seleziona un PDF o immagine."); return; }
    setStatus("OCR in corso nel browser...");
    byId("ocrText").textContent = "Lettura in corso...";
    try{
      var result = await Tesseract.recognize(currentFile, "ita+eng", { logger:function(m){ if(m.status) setStatus("OCR: "+m.status+(m.progress ? " "+Math.round(m.progress*100)+"%" : "")); }});
      currentOcrText = result.data && result.data.text || "";
      byId("ocrText").textContent = currentOcrText || "Nessun testo letto.";
      applyOcrSuggestions(currentOcrText);
      setStatus("OCR completato. Verifica e conferma i dati prima del salvataggio.");
    }catch(e){ console.error(e); setStatus("OCR non riuscito. Puoi compilare manualmente i campi."); byId("ocrText").textContent = "Errore OCR: "+e.message; }
  }
  function applyOcrSuggestions(text){
    var plate = guessPlate(text); if(plate) byId("plateInput").value = plate;
    var amount = guessAmount(text); if(amount) byId("amountInput").value = amount;
    var date = guessDate(text); if(date) byId("dateInput").value = date;
    var due = guessDueDate(text); if(due) byId("dueDateInput").value = due;
    var num = guessNumber(text); if(num) byId("numberInput").value = num;
    var asset = matchAssetByPlate(byId("plateInput").value); if(asset) renderAssetSelect(entityId(asset));
  }
  function guessPlate(t){ var m = String(t||"").toUpperCase().match(/\b[A-Z]{2}\s?[0-9]{3}\s?[A-Z]{2}\b/); return m ? norm(m[0]) : ""; }
  function guessAmount(t){ var m = String(t||"").match(/(?:€|EUR|euro)?\s*([0-9]{1,4}[,.][0-9]{2})\s*(?:€|EUR|euro)?/i); return m ? m[1].replace(",", ".") : ""; }
  function toIsoDate(d,m,y){ y=String(y); if(y.length===2) y="20"+y; return y+"-"+String(m).padStart(2,"0")+"-"+String(d).padStart(2,"0"); }
  function guessDate(t){ var m=String(t||"").match(/\b([0-3]?\d)[\/\-.]([01]?\d)[\/\-.]((?:20)?\d{2})\b/); return m ? toIsoDate(m[1],m[2],m[3]) : ""; }
  function guessDueDate(t){ var lower=String(t||"").toLowerCase(); var idx=lower.search(/scadenza|pagamento entro|entro il/); if(idx<0) return ""; return guessDate(lower.slice(idx, idx+120)); }
  function guessNumber(t){ var m=String(t||"").match(/(?:verbale|n\.|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i); return m ? m[1] : ""; }

  async function saveDocument(){
    currentFile = byId("fileInput").files[0];
    if(!currentFile){ alert("Seleziona un file."); return; }
    if(!byId("confirmInput").checked){ alert("Conferma obbligatoria prima del salvataggio."); return; }
    var asset = selectedAsset();
    var type = byId("docType").value;
    var plate = norm(byId("plateInput").value || (asset ? plateOf(asset) : ""));
    var id = uid(type === "fine" ? "fine" : "doc");
    var ext = extensionOf(currentFile);
    var fileName = [todayIso(), plate || "NO-PLATE", id].join("_") + "." + ext;
    var filePath = "documents/" + folderForType(type) + "/" + fileName;

    var record = {
      id:id,
      type:type,
      typeLabel:typeLabel(type),
      status:byId("statusInput").value,
      plate:plate,
      deviceId:asset ? entityId(asset) : "",
      assetName:asset ? nameOf(asset) : "",
      documentDate:byId("dateInput").value || "",
      dueDate:byId("dueDateInput").value || "",
      amount:byId("amountInput").value ? Number(byId("amountInput").value) : null,
      documentNumber:byId("numberInput").value.trim(),
      issuer:byId("issuerInput").value.trim(),
      notes:byId("notesInput").value.trim(),
      filePath:filePath,
      originalFileName:currentFile.name,
      ocrText:currentOcrText,
      confirmedByUser:true,
      createdAt:new Date().toISOString(),
      modifiedAt:new Date().toISOString()
    };

    try{
      setStatus("Carico il file documento su GitHub...");
      await ghPutBinary(filePath, currentFile, "Upload "+typeLabel(type)+" "+id);
      setStatus("Aggiorno indice documents.json senza sovrascrivere i record esistenti...");
      var latest = await ghGetFile("data/documents.json");
      var list = latest.exists && latest.content ? JSON.parse(latest.content) : [];
      list.push(record);
      var res = await ghPutFile("data/documents.json", JSON.stringify(list,null,2)+"\n", "Add document record "+id, latest.sha || ghIndexSha);
      ghIndexSha = res.content && res.content.sha || null;
      documents = list;
      renderAll(); clearForm(false);
      setStatus("Documento salvato. Record totali in archivio: "+documents.length+".");
    }catch(e){ console.error(e); alert(e.message); setStatus("Salvataggio non riuscito: "+e.message); }
  }

  function clearForm(clearFile){
    if(clearFile !== false) byId("fileInput").value = "";
    ["plateInput","dateInput","dueDateInput","amountInput","numberInput","issuerInput","notesInput"].forEach(function(id){ byId(id).value=""; });
    byId("confirmInput").checked=false; byId("saveDocument").disabled=true; byId("ocrText").textContent="Nessun OCR eseguito."; currentOcrText=""; currentFile=null; renderAssetSelect();
  }

  function renderAll(){ renderSummary(); renderArchive(); renderAssetCards(); }
  function renderSummary(){
    var fines = documents.filter(function(d){ return d.type === "fine"; });
    var open = fines.filter(function(d){ return ["da_pagare","da_verificare","contestata"].indexOf(d.status) !== -1; });
    var amt = open.reduce(function(s,d){ return s + (Number(d.amount)||0); },0);
    byId("docCount").textContent = documents.length;
    byId("fineCount").textContent = fines.length;
    byId("openFineCount").textContent = open.length;
    byId("openAmount").textContent = amt ? amt.toLocaleString("it-IT",{style:"currency",currency:"EUR"}) : "€0";
  }
  function renderArchive(){
    var q = (byId("searchInput").value || "").toLowerCase();
    var rows = documents.filter(function(d){ return !q || JSON.stringify(d).toLowerCase().indexOf(q) !== -1; }).slice().sort(function(a,b){ return String(b.createdAt||"").localeCompare(String(a.createdAt||"")); });
    var tbody=byId("archiveTable");
    if(!rows.length){ tbody.innerHTML='<tr><td colspan="8"><div class="empty-state">Nessun documento trovato.</div></td></tr>'; return; }
    tbody.innerHTML = rows.map(function(d){
      var link = d.filePath ? '<a class="doc-link" href="'+esc(rawUrl(d.filePath))+'" target="_blank" rel="noopener">Apri</a>' : '-';
      return '<tr><td><span class="pill '+esc(d.type)+'">'+esc(typeLabel(d.type))+'</span></td><td>'+esc(d.assetName||'-')+'</td><td>'+esc(d.plate||'-')+'</td><td>'+esc(d.documentDate||'-')+'</td><td>'+esc(d.dueDate||'-')+'</td><td>'+esc(d.amount != null ? Number(d.amount).toLocaleString("it-IT",{style:"currency",currency:"EUR"}) : '-')+'</td><td>'+esc(d.status||'-')+'</td><td>'+link+'</td></tr>';
    }).join("");
  }
  function renderAssetCards(){
    var groups = {};
    documents.forEach(function(d){ var key = d.deviceId || d.plate || "non_associati"; if(!groups[key]) groups[key]=[]; groups[key].push(d); });
    var keys = Object.keys(groups);
    var box = byId("assetCards");
    if(!keys.length){ box.innerHTML='<div class="empty-state">Nessun documento archiviato.</div>'; return; }
    box.innerHTML = keys.map(function(k){
      var list = groups[k]; var first=list[0];
      return '<div class="asset-card"><h3>'+esc(first.assetName || first.plate || 'Non associati')+'</h3><div class="meta">'+esc(first.plate || '-')+' · '+list.length+' documenti</div><ul>'+list.slice(0,6).map(function(d){return '<li>'+esc(typeLabel(d.type))+' · '+esc(d.documentDate || d.createdAt.slice(0,10))+' · <a class="doc-link" target="_blank" rel="noopener" href="'+esc(rawUrl(d.filePath))+'">Apri</a></li>';}).join('')+'</ul></div>';
    }).join("");
  }

  function wire(){
    loadLocalCfg();
    byId("saveConfig").addEventListener("click", saveLocalCfg);
    byId("loadArchive").addEventListener("click", function(){ loadArchive().catch(function(e){ alert(e.message); setStatus(e.message); }); });
    byId("runOcr").addEventListener("click", runOcr);
    byId("clearForm").addEventListener("click", function(){ clearForm(true); });
    byId("saveDocument").addEventListener("click", saveDocument);
    byId("confirmInput").addEventListener("change", function(){ byId("saveDocument").disabled = !this.checked; });
    byId("plateInput").addEventListener("input", function(){ var a=matchAssetByPlate(this.value); if(a) renderAssetSelect(entityId(a)); });
    byId("searchInput").addEventListener("input", renderArchive);
    document.querySelectorAll(".tab").forEach(function(tab){ tab.addEventListener("click", function(){ document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")}); tab.classList.add("active"); document.querySelectorAll(".tab-panel").forEach(function(p){p.classList.remove("active")}); byId("tab-"+tab.getAttribute("data-tab")).classList.add("active"); }); });
  }

  if(!window.geotab || !window.geotab.addin){ byId("localWarning").className = "notice warning"; wire(); }
  else{
    window.geotab.addin[ADDIN_NAMESPACE] = function(){ return { initialize:function(api,state,callback){ apiRef=api; pageState=state; wire(); loadAssets(); if(callback) callback(); }, focus:function(api,state){ apiRef=api; pageState=state; }, blur:function(){} }; };
  }
}());
