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
  var editingDocumentId = null;

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
  function renderAssetOptions(selectId, selectedId){
    var sel = byId(selectId);
    if(!sel) return;
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
  function statusLabel(type, value){
    var cfg = fieldConfig(type);
    var found = (cfg.statuses || []).find(function(s){ return s[0] === value; });
    return found ? found[1] : (value || "-");
  }
  function folderForType(t){ return ({fine:"multe",registration:"libretti",insurance:"assicurazioni",leasing:"leasing",inspection:"revisioni",fuel_card:"carte-carburante",other:"altro"})[t] || "altro"; }
  function fieldConfig(t){
    var cfg = {
      fine: {
        labels:{date:"Data infrazione",due:"Scadenza pagamento",amount:"Importo multa",number:"Numero verbale",issuer:"Ente emittente",status:"Stato multa"},
        show:{plate:true,asset:true,date:true,due:true,amount:true,number:true,issuer:true,notes:true,status:true},
        statuses:[["da_verificare","Da verificare"],["da_pagare","Da pagare"],["pagata","Pagata"],["contestata","Contestata"],["archiviata","Archiviata"]]
      },
      registration: {
        labels:{date:"Data immatricolazione",number:"Numero libretto",issuer:"Ufficio / fonte",status:"Stato documento"},
        show:{plate:true,asset:true,date:true,due:false,amount:false,number:true,issuer:false,notes:true,status:true},
        statuses:[["valido","Valido"],["da_verificare","Da verificare"],["archiviato","Archiviato"]]
      },
      insurance: {
        labels:{date:"Decorrenza",due:"Scadenza polizza",amount:"Premio / importo",number:"Numero polizza",issuer:"Compagnia assicurativa",status:"Stato polizza"},
        show:{plate:true,asset:true,date:true,due:true,amount:true,number:true,issuer:true,notes:true,status:true},
        statuses:[["attiva","Attiva"],["in_scadenza","In scadenza"],["scaduta","Scaduta"],["da_verificare","Da verificare"],["archiviata","Archiviata"]]
      },
      leasing: {
        labels:{date:"Decorrenza",due:"Scadenza contratto",amount:"Canone / importo",number:"Numero contratto",issuer:"Società leasing / noleggio",status:"Stato contratto"},
        show:{plate:true,asset:true,date:true,due:true,amount:true,number:true,issuer:true,notes:true,status:true},
        statuses:[["attivo","Attivo"],["in_scadenza","In scadenza"],["chiuso","Chiuso"],["da_verificare","Da verificare"],["archiviato","Archiviato"]]
      },
      inspection: {
        labels:{date:"Data revisione",due:"Prossima scadenza",number:"Numero pratica",issuer:"Centro revisione",status:"Stato revisione"},
        show:{plate:true,asset:true,date:true,due:true,amount:false,number:true,issuer:true,notes:true,status:true},
        statuses:[["regolare","Regolare"],["in_scadenza","In scadenza"],["scaduta","Scaduta"],["da_verificare","Da verificare"],["archiviata","Archiviata"]]
      },
      fuel_card: {
        labels:{date:"Data emissione",due:"Scadenza carta",number:"Numero carta",issuer:"Emittente",status:"Stato carta"},
        show:{plate:true,asset:true,date:true,due:true,amount:false,number:true,issuer:true,notes:true,status:true},
        statuses:[["attiva","Attiva"],["bloccata","Bloccata"],["scaduta","Scaduta"],["da_verificare","Da verificare"],["archiviata","Archiviata"]]
      },
      other: {
        labels:{date:"Data documento",due:"Scadenza",amount:"Importo",number:"Numero documento",issuer:"Ente / fornitore",status:"Stato documento"},
        show:{plate:true,asset:true,date:true,due:true,amount:true,number:true,issuer:true,notes:true,status:true},
        statuses:[["da_verificare","Da verificare"],["valido","Valido"],["archiviato","Archiviato"]]
      }
    };
    return cfg[t] || cfg.other;
  }
  function setVisible(id, visible){ var el = byId(id); if(el) el.classList.toggle("hidden", !visible); }
  function applyDocumentType(){
    var t = byId("docType").value;
    var cfg = fieldConfig(t);
    var map = {date:"dateLabel", due:"dueDateLabel", amount:"amountLabel", number:"numberLabel", issuer:"issuerLabel", status:"statusLabel"};
    Object.keys(map).forEach(function(k){ if(cfg.labels[k] && byId(map[k])) byId(map[k]).textContent = cfg.labels[k]; });
    setVisible("plateField", cfg.show.plate); setVisible("assetField", cfg.show.asset); setVisible("dateField", cfg.show.date); setVisible("dueDateField", cfg.show.due); setVisible("amountField", cfg.show.amount); setVisible("numberField", cfg.show.number); setVisible("issuerField", cfg.show.issuer); setVisible("notesField", cfg.show.notes); setVisible("statusField", cfg.show.status);
    var prev = byId("statusInput").value;
    byId("statusInput").innerHTML = cfg.statuses.map(function(s){ return '<option value="'+esc(s[0])+'">'+esc(s[1])+'</option>'; }).join("");
    if(cfg.statuses.some(function(s){ return s[0] === prev; })) byId("statusInput").value = prev;
  }
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
    var t = byId("docType").value;
    var plate = guessPlate(text); if(plate) byId("plateInput").value = plate;
    var date = guessDate(text); if(date) byId("dateInput").value = date;

    if(["fine","insurance","leasing","other"].indexOf(t) !== -1){
      var amount = guessAmount(text); if(amount) byId("amountInput").value = amount;
    }
    if(["fine","insurance","leasing","inspection","fuel_card","registration","other"].indexOf(t) !== -1){
      var num = guessNumberByType(text, t); if(num) byId("numberInput").value = num;
    }
    if(["fine","insurance","leasing","inspection","fuel_card","other"].indexOf(t) !== -1){
      var due = guessDueDateByType(text, t); if(due) byId("dueDateInput").value = due;
    }
    if(["fine","insurance","leasing","inspection","fuel_card","other"].indexOf(t) !== -1){
      var issuer = guessIssuerByType(text, t); if(issuer) byId("issuerInput").value = issuer;
    }

    var asset = matchAssetByPlate(byId("plateInput").value); if(asset) renderAssetSelect(entityId(asset));
  }
  function guessPlate(t){ var m = String(t||"").toUpperCase().match(/\b[A-Z]{2}\s?[0-9]{3}\s?[A-Z]{2}\b/); return m ? norm(m[0]) : ""; }
  function guessAmount(t){ var m = String(t||"").match(/(?:€|EUR|euro)?\s*([0-9]{1,4}[,.][0-9]{2})\s*(?:€|EUR|euro)?/i); return m ? m[1].replace(",", ".") : ""; }
  function toIsoDate(d,m,y){ y=String(y); if(y.length===2) y="20"+y; return y+"-"+String(m).padStart(2,"0")+"-"+String(d).padStart(2,"0"); }
  function guessDate(t){ var m=String(t||"").match(/\b([0-3]?\d)[\/\-.]([01]?\d)[\/\-.]((?:20)?\d{2})\b/); return m ? toIsoDate(m[1],m[2],m[3]) : ""; }
  function guessDueDate(t){ var lower=String(t||"").toLowerCase(); var idx=lower.search(/scadenza|pagamento entro|entro il/); if(idx<0) return ""; return guessDate(lower.slice(idx, idx+120)); }
  function guessNumber(t){ var m=String(t||"").match(/(?:verbale|n\.|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i); return m ? m[1] : ""; }
  function guessNumberByType(t, type){
    var text = String(t||"");
    var patterns = {
      fine: /(?:verbale|accertamento|n\.|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i,
      insurance: /(?:polizza|contratto|n\.|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i,
      leasing: /(?:contratto|noleggio|leasing|n\.|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i,
      registration: /(?:libretto|carta di circolazione|telaio|vin)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i,
      inspection: /(?:revisione|pratica|n\.|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i,
      fuel_card: /(?:carta|card|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i
    };
    var m = text.match(patterns[type] || /(?:documento|n\.|numero)\s*[:#-]?\s*([A-Z0-9\/-]{4,})/i);
    return m ? m[1] : guessNumber(text);
  }
  function guessDueDateByType(t, type){
    var lower=String(t||"").toLowerCase();
    var keywords = {
      fine:/scadenza|pagamento entro|entro il|termine pagamento/,
      insurance:/scadenza|scade|valid[ao] fino|copertura fino/,
      leasing:/scadenza|fine contratto|termine contratto|fino al/,
      inspection:/prossima revisione|scadenza|revisione entro/,
      fuel_card:/scadenza|valida fino|valid thru/,
      other:/scadenza|fino al|entro il/
    };
    var re = keywords[type] || keywords.other;
    var idx=lower.search(re); if(idx<0) return ""; return guessDate(lower.slice(idx, idx+160));
  }
  function guessIssuerByType(t, type){
    var lines = String(t||"").split(/\r?\n/).map(function(x){return x.trim();}).filter(Boolean);
    if(!lines.length) return "";
    if(type === "insurance"){
      var l = lines.find(function(x){ return /assicur|insurance|groupama|generali|allianz|unipol|reale|axa|zurich|vittoria|cattolica/i.test(x); });
      return l ? l.slice(0,80) : "";
    }
    if(type === "fine"){
      var l2 = lines.find(function(x){ return /comune|polizia|municipale|prefettura|carabinieri|verbale/i.test(x); });
      return l2 ? l2.slice(0,80) : "";
    }
    if(type === "leasing"){
      var l3 = lines.find(function(x){ return /leasing|noleggio|lease|rent|arval|leaseplan|ald|leasys/i.test(x); });
      return l3 ? l3.slice(0,80) : "";
    }
    return "";
  }

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
      company:type === "insurance" ? byId("issuerInput").value.trim() : "",
      policyNumber:type === "insurance" ? byId("numberInput").value.trim() : "",
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
    byId("confirmInput").checked=false; byId("saveDocument").disabled=true; byId("ocrText").textContent="Nessun OCR eseguito."; currentOcrText=""; currentFile=null; renderAssetSelect(); applyDocumentType();
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
    if(!rows.length){ tbody.innerHTML='<tr><td colspan="9"><div class="empty-state">Nessun documento trovato.</div></td></tr>'; return; }
    tbody.innerHTML = rows.map(function(d){
      var link = d.filePath ? '<a class="doc-link" href="'+esc(rawUrl(d.filePath))+'" target="_blank" rel="noopener">Apri</a>' : '-';
      var amount = d.amount != null ? Number(d.amount).toLocaleString("it-IT",{style:"currency",currency:"EUR"}) : '-';
      return '<tr><td><span class="pill '+esc(d.type)+'">'+esc(typeLabel(d.type))+'</span></td><td>'+esc(d.assetName||'-')+'</td><td>'+esc(d.plate||'-')+'</td><td>'+esc(d.documentDate||'-')+'</td><td>'+esc(d.dueDate||'-')+'</td><td>'+esc(amount)+'</td><td>'+esc(statusLabel(d.type, d.status))+'</td><td>'+link+'</td><td><button class="btn small js-edit-doc" data-id="'+esc(d.id)+'">Modifica</button></td></tr>';
    }).join("");
  }

  function openEditDocument(id){
    var d = documents.find(function(x){ return x.id === id; });
    if(!d){ alert("Documento non trovato in archivio."); return; }
    editingDocumentId = id;
    byId("editId").value = id;
    byId("editTypeLabel").value = typeLabel(d.type);
    byId("editPlateInput").value = d.plate || "";
    renderAssetOptions("editAssetSelect", d.deviceId || "");
    var cfg = fieldConfig(d.type);
    byId("editStatusInput").innerHTML = (cfg.statuses || []).map(function(s){ return '<option value="'+esc(s[0])+'">'+esc(s[1])+'</option>'; }).join("");
    byId("editStatusInput").value = d.status || ((cfg.statuses && cfg.statuses[0] && cfg.statuses[0][0]) || "");
    byId("editDateInput").value = d.documentDate || "";
    byId("editDueDateInput").value = d.dueDate || "";
    byId("editAmountInput").value = d.amount != null ? d.amount : "";
    byId("editNumberInput").value = d.documentNumber || d.policyNumber || "";
    byId("editIssuerInput").value = d.issuer || d.company || "";
    byId("editNotesInput").value = d.notes || "";
    byId("editPanel").classList.remove("hidden");
    byId("editPanel").scrollIntoView({behavior:"smooth", block:"start"});
  }

  function closeEditDocument(){
    editingDocumentId = null;
    byId("editPanel").classList.add("hidden");
  }

  function findAssetById(id){
    return assets.find(function(a){ return entityId(a) === id; }) || null;
  }

  async function saveEditDocument(){
    var id = byId("editId").value || editingDocumentId;
    if(!id){ alert("Nessun documento selezionato."); return; }
    try{
      setStatus("Aggiorno archivio su GitHub...");
      var latest = await ghGetFile("data/documents.json");
      var list = latest.exists && latest.content ? JSON.parse(latest.content) : documents.slice();
      var idx = list.findIndex(function(d){ return d.id === id; });
      if(idx < 0) throw new Error("Documento non trovato nel file documents.json aggiornato.");
      var d = list[idx];
      var asset = findAssetById(byId("editAssetSelect").value);
      d.status = byId("editStatusInput").value;
      d.plate = norm(byId("editPlateInput").value || (asset ? plateOf(asset) : d.plate));
      d.deviceId = asset ? entityId(asset) : byId("editAssetSelect").value || "";
      d.assetName = asset ? nameOf(asset) : d.assetName || "";
      d.documentDate = byId("editDateInput").value || "";
      d.dueDate = byId("editDueDateInput").value || "";
      d.amount = byId("editAmountInput").value ? Number(byId("editAmountInput").value) : null;
      d.documentNumber = byId("editNumberInput").value.trim();
      d.issuer = byId("editIssuerInput").value.trim();
      if(d.type === "insurance"){
        d.company = d.issuer;
        d.policyNumber = d.documentNumber;
      }
      d.notes = byId("editNotesInput").value.trim();
      d.modifiedAt = new Date().toISOString();
      list[idx] = d;
      var res = await ghPutFile("data/documents.json", JSON.stringify(list,null,2)+"\n", "Update document record "+id, latest.sha || ghIndexSha);
      ghIndexSha = res.content && res.content.sha || null;
      documents = list;
      renderAll();
      closeEditDocument();
      setStatus("Archivio aggiornato. Documento modificato: "+id+".");
    }catch(e){ console.error(e); alert(e.message); setStatus("Modifica non riuscita: "+e.message); }
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
    byId("docType").addEventListener("change", applyDocumentType);
    applyDocumentType();
    byId("runOcr").addEventListener("click", runOcr);
    byId("clearForm").addEventListener("click", function(){ clearForm(true); });
    byId("saveDocument").addEventListener("click", saveDocument);
    byId("confirmInput").addEventListener("change", function(){ byId("saveDocument").disabled = !this.checked; });
    byId("plateInput").addEventListener("input", function(){ var a=matchAssetByPlate(this.value); if(a) renderAssetSelect(entityId(a)); });
    byId("searchInput").addEventListener("input", renderArchive);
    byId("saveEdit").addEventListener("click", saveEditDocument);
    byId("cancelEdit").addEventListener("click", closeEditDocument);
    byId("editPlateInput").addEventListener("input", function(){ var a=matchAssetByPlate(this.value); if(a) renderAssetOptions("editAssetSelect", entityId(a)); });
    document.body.addEventListener("click", function(event){ var t = event.target; if(t && t.classList.contains("js-edit-doc")){ openEditDocument(t.getAttribute("data-id")); } });
    document.querySelectorAll(".tab").forEach(function(tab){ tab.addEventListener("click", function(){ document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")}); tab.classList.add("active"); document.querySelectorAll(".tab-panel").forEach(function(p){p.classList.remove("active")}); byId("tab-"+tab.getAttribute("data-tab")).classList.add("active"); }); });
  }

  if(!window.geotab || !window.geotab.addin){ byId("localWarning").className = "notice warning"; wire(); }
  else{
    window.geotab.addin[ADDIN_NAMESPACE] = function(){ return { initialize:function(api,state,callback){ apiRef=api; pageState=state; wire(); loadAssets(); if(callback) callback(); }, focus:function(api,state){ apiRef=api; pageState=state; }, blur:function(){} }; };
  }
}());
