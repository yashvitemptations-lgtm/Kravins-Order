// ═══════════════════════════════════════════════════════════════════════════
// SALES ORDER MANAGEMENT SYSTEM - Apps Script MVP
// Sheet ID: 174sQJSqaTZDcJFZWVtVMe6Df4D1f0Lg0h1RRkGc397c
// ═══════════════════════════════════════════════════════════════════════════

var SHEET_ID = '174sQJSqaTZDcJFZWVtVMe6Df4D1f0Lg0h1RRkGc397c';

var TAB_ORDERS      = 'Orders';
var TAB_PARTIES     = 'New_Parties';
var TAB_PRODUCTS    = 'Master_Products';
var TAB_SALESMEN    = 'Master_Salesmen';
var TAB_DISTRIBUTORS= 'Master_Distributors';
var TAB_BUYERS      = 'Master_Buyers';

var HDR_ORDERS = [
  'Order ID','Date','Time','Status','Salesman ID','Salesman Name',
  'Buyer ID','Buyer Name','Buyer Phone','Buyer Address',
  'Product Code','Product Name','Qty','Unit Price','Line Total',
  'Commission %','Commission Amt','Notes','Photo URL','Distributor'
];

var HDR_PARTIES = [
  'Buyer ID','Party Name','Contact Person','Phone','Address',
  'Route','Territory','GST','Terms','Credit Limit','Outstanding',
  'Added By','Added At','Status','Latitude','Longitude','Photo URL',
  'WhatsApp','Email','Notes','Day'
];

var HDR_PRODUCTS = [
  'Code','Name','Category','Price','Commission %','Active'
];

var HDR_SALESMEN = [
  'ID','Name','Phone','Territory','Target','Commission Enabled','Distributor','Active','PIN',
  'WhatsApp','Commission %','Pay Cycle','Address','City','State',
  'PAN','Aadhaar','Bank','Account No','IFSC','Account Name','Notes',
  'Doc PAN','Doc ID','Doc Address','Doc Bank','Doc Photo',
  'Shops Target','Parties Target','Prod Mix Targets','Daily Target','Daily Shops Target','Joining Date','Visit Target'
];

var HDR_DISTRIBUTORS = [
  'ID','Name','Phone','WhatsApp','Territory','Active','PIN'
];

var HDR_BUYERS = [
  'ID','Name','Contact','Phone','Address','Route','Territory',
  'GST','Terms','Credit Limit','Outstanding','Active'
];

// ════════════════════════════════════════════════════════════════════════
// ENTRY POINTS
// ════════════════════════════════════════════════════════════════════════

function doGet(e) {
  var p = e.parameter;
  var r;
  try {
    if      (p.action === 'dashboard')    r = getDashboard(p.smId, p.month, p.year);
    else if (p.action === 'orders')       r = getOrders(p.smId, p.distId);
    else if (p.action === 'products')     r = getProducts();
    else if (p.action === 'salesmen')     r = getSalesmen();
    else if (p.action === 'distributors') r = getDistributors();
    else if (p.action === 'buyers')       r = getBuyers(p.smId);
    else if (p.action === 'parties')      r = getParties(p.smId);
    else if (p.action === 'partyAccess')   r = getPartyAccess();
    else if (p.action === 'verifyPin')     r = verifyPin(p.id, p.pin, p.role);
    else if (p.action === 'editRequests')    r = getEditRequests();
    else if (p.action === 'migrateDistAlloc')  r = migrateDistAlloc();
    else if (p.action === 'distAlloc')          r = {ok:true, partyIds:getDistPartyAlloc(p.distId)};
    else if (p.action === 'visits')             r = getVisits(p.smId, p.distId, p.from||p.dateFrom, p.to||p.dateTo);
    else if (p.action === 'visitStats')         r = getVisitStats(p.smId);
    else r = {ok:false, error:'Unknown action'};
  } catch(e) {
    r = {ok:false, error:e.toString()};
  }
  return out(r);
}

function doPost(e) {
  var r;
  try {
    var b = JSON.parse(e.postData.contents);
    if      (b.action === 'submitOrder')    r = saveOrder(b.order);
    else if (b.action === 'addParty')       r = saveParty(b.party);
    else if (b.action === 'addProduct')     r = saveProduct(b.product);
    else if (b.action === 'addSalesman')    r = saveSalesman(b.salesman);
    else if (b.action === 'addDistributor') r = saveDistributor(b.distributor);
    else if (b.action === 'addBuyer')       r = saveBuyer(b.buyer);
    else if (b.action === 'updateStatus')   r = updateStatus(b.orderId, b.status);
    else if (b.action === 'editOrder')      r = editOrder(b.orderId, b.changes, b.actor);
    else if (b.action === 'cancelOrder')    r = cancelOrder(b.orderId, b.reason, b.actor);
    else if (b.action === 'transferParty')  r = transferParty(b.partyId, b.fromSm, b.toSm, b.adminNote);
    else if (b.action === 'saveProduct')    r = saveProductMaster(b.product);
    else if (b.action === 'savePartyAccess')      r = savePartyAccess(b.access);
    else if (b.action === 'bulkImportParties')    r = bulkImportParties(b.parties, b.smId);
    else if (b.action === 'savePartyAllocations') r = savePartyAllocations(b.smId, b.partyIds);
    else if (b.action === 'updatePartyStatus')   r = updatePartyStatus(b.partyId, b.status);
    else if (b.action === 'saveVisit')           r = saveVisit(b.visit||b);
    else if (b.action === 'saveVisit')            r = saveVisit(b.visit||b);
    else if (b.action === 'deactivateSalesman')  r = deactivateSalesman(b.smId, b.active);
    else if (b.action === 'deactivateDistributor')    r = deactivateDistributor(b.distId, b.active);
    else if (b.action === 'updatePartyFull')           r = updatePartyFull(b.party);
    else if (b.action === 'submitEditRequest')         r = submitEditRequest(b);
    else if (b.action === 'resolveEditRequest')        r = resolveEditRequest(b.reqId, b.resolution);
    else if (b.action === 'saveDistPartyAllocations')  r = saveDistPartyAllocations(b.distId, b.partyIds, b);
    else if (b.action === 'saveVisit')                  r = saveVisit(b.visit);
    else r = {ok:false, error:'Unknown action'};
  } catch(e) {
    r = {ok:false, error:e.toString()};
  }
  return out(r);
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════
// SAVE ORDER
// ════════════════════════════════════════════════════════════════════════

function saveOrder(order) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_ORDERS, HDR_ORDERS);
  var now = new Date();

  var indianDate = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  var timeStr    = Utilities.formatDate(now, 'Asia/Kolkata', 'HH:mm:ss');

  var items = order.items || [];
  for (var i=0; i<items.length; i++) {
    var it = items[i];
    sheet.appendRow([
      order.oid, indianDate, timeStr, 'Pending',
      String(order.smId||'').trim(), String(order.smName||'').trim(),
      order.bid, order.bname, order.bphone || '', order.baddress || '',
      it.code, it.name, it.qty, it.price, it.total,
      it.comm || 0, it.commAmt || 0,
      order.notes || '', order.photoUrl || '', order.distributor || ''
    ]);
  }

  colorRows(sheet);
  return {ok:true, oid:order.oid};
}

// ════════════════════════════════════════════════════════════════════════
// EDIT ORDER (notes + status, within 30-min window enforced by HTML)
// ════════════════════════════════════════════════════════════════════════

function editOrder(orderId, changes, actor) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_ORDERS);
  if (!sheet) return {ok:false, error:'Orders sheet not found'};

  var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, HDR_ORDERS.length).getValues();
  var updated = 0;
  for (var i=0; i<rows.length; i++) {
    if (rows[i][0] === orderId) {
      var rowNum = i + 2;
      if (changes.status !== undefined) sheet.getRange(rowNum, 4).setValue(changes.status);
      if (changes.notes  !== undefined) sheet.getRange(rowNum, 18).setValue(changes.notes);
      updated++;
    }
  }
  return {ok:true, updated:updated};
}

// ════════════════════════════════════════════════════════════════════════
// CANCEL ORDER
// ════════════════════════════════════════════════════════════════════════

function cancelOrder(orderId, reason, actor) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_ORDERS);
  if (!sheet) return {ok:false, error:'Orders sheet not found'};

  var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, HDR_ORDERS.length).getValues();
  var updated = 0;
  for (var i=0; i<rows.length; i++) {
    if (rows[i][0] === orderId) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, 4).setValue('Cancelled');
      if (reason) {
        var existingNotes = rows[i][17] || '';
        sheet.getRange(rowNum, 18).setValue(
          (existingNotes ? existingNotes + ' | ' : '') + 'CANCELLED: ' + reason
        );
      }
      updated++;
    }
  }
  return {ok:true, updated:updated};
}

// ════════════════════════════════════════════════════════════════════════
// TRANSFER PARTY  (admin only)
//
// Rules:
//  - Changes the AddedBy column for the party in New_Parties sheet
//  - Orders already placed (smId on Orders rows) are NOT touched
//  - So the original salesman keeps credit for all orders they placed
//  - New salesman gets credit only for orders they place AFTER transfer
//  - A transfer log note is written to the Status column
// ════════════════════════════════════════════════════════════════════════

function transferParty(partyId, fromSm, toSm, adminNote) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_PARTIES);
  if (!sheet) return {ok:false, error:'New_Parties sheet not found'};

  var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, HDR_PARTIES.length).getValues();
  var found = false;
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]).trim() === String(partyId).trim()) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, 12).setValue(toSm);   // AddedBy → new SM
      var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm');
      var note = 'Transferred ' + fromSm + '→' + toSm + ' on ' + now;
      if (adminNote) note += ' (' + adminNote + ')';
      sheet.getRange(rowNum, 14).setValue(note);    // Status → transfer log
      found = true;
    }
  }
  if (!found) return {ok:false, error:'Party not found: ' + partyId};
  return {ok:true, partyId:partyId, from:fromSm, to:toSm};
}

// ════════════════════════════════════════════════════════════════════════
// SAVE PARTY
// ════════════════════════════════════════════════════════════════════════

function saveParty(p) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_PARTIES, HDR_PARTIES);
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm');

  // Generate unique ID server-side to avoid duplicates from stale client cache
  var existingIds = {};
  if (sheet.getLastRow() >= 2) {
    var idCol = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0; i<idCol.length; i++) {
      var eid = String(idCol[i][0]||'').trim();
      if (eid) existingIds[eid] = true;
    }
  }
  // Check if client-provided ID is already taken
  var pid = String(p.id||'').trim();
  if (!pid || existingIds[pid]) {
    // Generate new unique ID
    var nums = Object.keys(existingIds)
      .map(function(id){ return parseInt(id.replace(/[^0-9]/g,''))||0; })
      .filter(function(n){ return n>0; });
    var maxNum = nums.length > 0 ? Math.max.apply(null, nums) : 0;
    var prefix = pid ? pid.replace(/[0-9]+$/,'') : 'BP';
    pid = prefix + String(maxNum+1).padStart(4,'0');
    // Ensure still unique
    while (existingIds[pid]) {
      maxNum++;
      pid = prefix + String(maxNum+1).padStart(4,'0');
    }
  }

  sheet.appendRow([
    pid, p.name, p.contact||'', p.phone||'', p.address||'',
    p.route||'', p.territory||'', p.gst||'', p.terms||'',
    p.credit||0, p.outstanding||0, p.addedBy||'', now,
    p.status||'Active', p.lat||'', p.lng||'', p.photo||'',
    p.whatsapp||'', p.email||'', p.notes||'', p.day||''
  ]);
  return {ok:true, bid:pid};
}

// ════════════════════════════════════════════════════════════════════════
// MASTER DATA - PRODUCTS
// ════════════════════════════════════════════════════════════════════════

function getProducts() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_PRODUCTS);
  if (!sheet) return {ok:true, products:getDefaultProducts()};
  if (sheet.getLastRow() < 2) return {ok:true, products:getDefaultProducts()};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_PRODUCTS.length).getValues();
  var prods = [];
  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var isActive = r[5]!=='No';
    prods.push({
      code:r[0], name:r[1], category:r[2]||'',
      price:r[3], comm:r[4]||0, active:isActive
    });
  }
  // For salesman app: only return active products
  // getProducts is called without params; admin product master handles inactive display separately
  var activeProds = prods.filter(function(p){ return p.active !== false; });
  return {ok:true, products:activeProds.length ? activeProds : (prods.length ? prods : getDefaultProducts()), allProducts:prods};
}

function saveProduct(p) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_PRODUCTS, HDR_PRODUCTS);
  sheet.appendRow([p.code, p.name, p.category||'', p.price, p.comm||0, 'Yes']);
  return {ok:true};
}

function getDefaultProducts() {
  return [
    {code:'P001',name:'Crazy Poms Jar',price:150,comm:0.05,active:true},
    {code:'P002',name:'Tams Jar',price:150,comm:0.05,active:true},
    {code:'P003',name:'Crazy Poms 250g Pouch',price:50,comm:0.04,active:true},
    {code:'P004',name:'Tams 250g Pouch',price:50,comm:0.04,active:true},
    {code:'P005',name:'Crazy Poms Display Box (5pc)',price:240,comm:0.06,active:true},
    {code:'P006',name:'Tams Display Box (5pc)',price:240,comm:0.06,active:true},
    {code:'P007',name:'Crazy Poms Display Box (50pc)',price:350,comm:0.07,active:true},
    {code:'P008',name:'Tams Display Box (50pc)',price:350,comm:0.07,active:true},
    {code:'P009',name:'Mix Flavour Dangler (72pc)',price:500,comm:0.08,active:true},
    {code:'P010',name:'Mix Flavour Dangler (24pc)',price:180,comm:0.06,active:true},
    {code:'P011',name:'Mix Flavour Display Box (50pc)',price:360,comm:0.07,active:true}
  ];
}

// ════════════════════════════════════════════════════════════════════════
// MASTER DATA - SALESMEN
// ════════════════════════════════════════════════════════════════════════

function getSalesmen() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_SALESMEN);
  if (!sheet) return {ok:true, salesmen:getDefaultSalesmen()};
  if (sheet.getLastRow() < 2) return {ok:true, salesmen:getDefaultSalesmen()};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_SALESMEN.length).getValues();
  var sms = [];
  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    sms.push({
      id:r[0], name:r[1], phone:r[2]||'', territory:r[3]||'',
      target:r[4]||50000, commEnabled:r[5]!=='No',
      distributor:r[6]||'', active:r[7]!=='No',
      pin:String(r[8]||'').trim(),
      whatsapp:r[9]||'', commPct:r[10]||0, payCycle:r[11]||'Monthly',
      address:r[12]||'', city:r[13]||'', state:r[14]||'',
      pan:r[15]||'', aadhaar:r[16]||'', bank:r[17]||'',
      accountNo:r[18]||'', ifsc:r[19]||'', accountName:r[20]||'',
      notes:r[21]||'',
      docPan:r[22]||'', docId:r[23]||'', docAddress:r[24]||'',
      docBank:r[25]||'', docPhoto:r[26]||'',
      shopsTarget:r[27]||0, partiesTarget:r[28]||0,
      prodMixTargets:(function(){ try{return JSON.parse(r[29]||'{}');}catch(e){return {};} })(),
      dailyTarget:r[30]||0, dailyShopsTarget:r[31]||0,
      joiningDate:String(r[32]||'').trim(),
      visitTarget:r[33]||50
    });
  }
  return {ok:true, salesmen:sms.length ? sms : getDefaultSalesmen()};
}

function saveSalesman(s) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_SALESMEN, HDR_SALESMEN);
  var rowData = [
    s.id, s.name, s.phone||'', s.territory||'', s.target||50000,
    s.commEnabled ? 'Yes' : 'No', s.distributor||'',
    s.active===false ? 'No' : 'Yes',
    s.pin||'',
    s.whatsapp||'', s.commPct||0, s.payCycle||'Monthly',
    s.address||'', s.city||'', s.state||'',
    s.pan||'', s.aadhaar||'', s.bank||'', s.accountNo||'', s.ifsc||'', s.accountName||'',
    s.notes||'',
    s.docPan||'', s.docId||'', s.docAddress||'', s.docBank||'', s.docPhoto||'',
    s.shopsTarget||0, s.partiesTarget||0,
    s.prodMixTargets ? JSON.stringify(s.prodMixTargets) : '',
    s.dailyTarget||0, s.dailyShopsTarget||0, s.joiningDate||'', s.visitTarget||50
  ];
  // Update existing row if ID found, else append
  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0; i<rows.length; i++) {
      if (String(rows[i][0]).trim() === String(s.id).trim()) {
        sheet.getRange(i+2, 1, 1, HDR_SALESMEN.length).setValues([rowData]);
        return {ok:true, action:'updated'};
      }
    }
  }
  sheet.appendRow(rowData);
  return {ok:true, action:'created'};
}

function getDefaultSalesmen() {
  return [
    {id:'SM001',name:'Sunil Bhai',territory:'North',target:200000,commEnabled:true,distributor:'D001',active:true},
    {id:'SM002',name:'JK Bhai',territory:'South',target:50000,commEnabled:true,distributor:'D001',active:true},
    {id:'SM003',name:'Satish',territory:'East',target:50000,commEnabled:true,distributor:'D001',active:true},
    {id:'SM004',name:'Ajay',territory:'West',target:50000,commEnabled:false,distributor:'D001',active:true},
    {id:'SM005',name:'Jitu Bhai',territory:'Central',target:50000,commEnabled:true,distributor:'D001',active:true}
  ];
}

// ════════════════════════════════════════════════════════════════════════
// MASTER DATA - DISTRIBUTORS
// ════════════════════════════════════════════════════════════════════════

function getDistributors() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_DISTRIBUTORS);
  if (!sheet) return {ok:true, distributors:getDefaultDistributors()};
  if (sheet.getLastRow() < 2) return {ok:true, distributors:getDefaultDistributors()};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_DISTRIBUTORS.length).getValues();
  var dists = [];
  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    dists.push({
      id:r[0], name:r[1], phone:r[2]||'', whatsapp:r[3]||'',
      territory:r[4]||'', active:r[5]!=='No',
      pin:String(r[6]||'').trim()
    });
  }
  return {ok:true, distributors:dists.length ? dists : getDefaultDistributors()};
}

function saveDistributor(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_DISTRIBUTORS, HDR_DISTRIBUTORS);
  var rowData = [
    d.id, d.name, d.phone||'', d.whatsapp||'', d.territory||'',
    d.active===false ? 'No' : 'Yes',
    d.pin||''
  ];
  // Update existing row if ID found, else append
  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0; i<rows.length; i++) {
      if (String(rows[i][0]).trim() === String(d.id).trim()) {
        sheet.getRange(i+2, 1, 1, HDR_DISTRIBUTORS.length).setValues([rowData]);
        return {ok:true, action:'updated'};
      }
    }
  }
  sheet.appendRow(rowData);
  return {ok:true, action:'created'};
}

function getDefaultDistributors() {
  return [
    {id:'D001',name:'SKP Beingwell Solutions LLP',phone:'7490985034',whatsapp:'7490985034',territory:'New Ahmedabad',active:true}
  ];
}

// ════════════════════════════════════════════════════════════════════════
// MASTER DATA - BUYERS (Master_Buyers sheet — admin-managed list)
// No smId filter here; buyers are shared across all salesmen
// ════════════════════════════════════════════════════════════════════════

function getBuyers(smId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_BUYERS);
  if (!sheet) return {ok:true, buyers:[]};
  if (sheet.getLastRow() < 2) return {ok:true, buyers:[]};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_BUYERS.length).getValues();
  var buyers = [];
  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    buyers.push({
      id:r[0], name:r[1], contact:r[2]||'', phone:r[3]||'',
      address:r[4]||'', route:r[5]||'', territory:r[6]||'',
      gst:r[7]||'', terms:r[8]||'', credit:r[9]||0,
      outstanding:r[10]||0, active:r[11]!=='No'
    });
  }
  return {ok:true, buyers:buyers};
}

function saveBuyer(b) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_BUYERS, HDR_BUYERS);
  sheet.appendRow([
    b.id, b.name, b.contact||'', b.phone||'', b.address||'',
    b.route||'', b.territory||'', b.gst||'', b.terms||'',
    b.credit||0, 0, 'Yes'
  ]);
  return {ok:true};
}

// ════════════════════════════════════════════════════════════════════════
// GET PARTIES  (New_Parties sheet — salesman-added parties)
//
// smId passed  → return ONLY parties AddedBy that salesman (salesman app)
// smId omitted → return ALL parties with addedBy info (admin app)
// ════════════════════════════════════════════════════════════════════════

function getParties(smId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_PARTIES);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, parties:[]};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_PARTIES.length).getValues();
  var parties = [];
  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    if (smId && String(r[11]).trim() !== String(smId).trim()) continue;
    parties.push({
      id       : String(r[0]).trim(),
      name     : String(r[1]).trim(),
      contact  : String(r[2]||'').trim(),
      phone    : String(r[3]||'').trim(),
      address  : String(r[4]||'').trim(),
      route    : String(r[5]||'').trim(),
      territory: String(r[6]||'').trim(),
      gst      : String(r[7]||'').trim(),
      terms    : String(r[8]||'').trim(),
      credit   : r[9]||0,
      outstanding: r[10]||0,
      addedBy  : String(r[11]||'').trim(),
      at       : String(r[12]||'').trim(),
      status   : String(r[13]||'').trim(),
      lat      : String(r[14]||'').trim(),
      lng      : String(r[15]||'').trim(),
      photo    : String(r[16]||'').trim(),
      whatsapp : String(r[17]||'').trim(),
      email    : String(r[18]||'').trim(),
      notes    : String(r[19]||'').trim(),
      day      : String(r[20]||'').trim()
    });
  }
  // Merge dist allocations from separate Dist_Alloc sheet
  var distAllocMap = getAllDistAlloc();
  parties.forEach(function(p){ p.distAlloc = distAllocMap[p.id] || ''; });

  return {ok:true, parties:parties};
}

// ════════════════════════════════════════════════════════════════════════
// GET ORDERS
// smId  → salesman app: only their orders
// distId→ distributor app: only their distributor's orders
// none  → admin: all orders
// ════════════════════════════════════════════════════════════════════════

function getOrders(smId, distId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, orders:[]};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_ORDERS.length).getValues();
  var map = {};

  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    var oid = r[0];
    if (!oid) continue;
    if (smId   && String(r[4]||'').trim()  !== String(smId).trim())   continue;
    if (distId && String(r[19]||'').trim() !== String(distId).trim()) continue;

    if (!map[oid]) {
      map[oid] = {
        oid:r[0], date:r[1], time:r[2], status:r[3],
        smId:r[4], smName:r[5],
        bid:r[6], bname:r[7], bphone:r[8], baddress:r[9],
        notes:r[17], photoUrl:r[18], distributor:r[19],
        gt:0, tc:0, items:[], at:r[2]
      };
    }

    var lt = parseFloat(r[14]) || 0;
    var ca = parseFloat(r[16]) || 0;
    map[oid].gt += lt;
    map[oid].tc += ca;
    map[oid].items.push({
      code:r[10], name:r[11], qty:r[12],
      price:r[13], total:lt, comm:r[15], commAmt:ca
    });
  }

  var list = [];
  for (var k in map) list.push(map[k]);
  list.sort(function(a,b){ return b.oid>a.oid?1:-1; });

  return {ok:true, orders:list};
}

// ════════════════════════════════════════════════════════════════════════
// GET DASHBOARD
// All KPIs are based on orders placed BY that salesman (smId filter).
// newParties count & array are filtered to parties ADDED BY that salesman.
// Admin (no smId) sees company-wide totals.
// ════════════════════════════════════════════════════════════════════════

function getDashboard(smId, month, year) {
  var allOrdersList = getOrders().orders;
  var orders = smId
    ? allOrdersList.filter(function(o){ return o.smId === smId; })
    : allOrdersList;

  var salesmen   = getSalesmen().salesmen;
  var allParties = getParties().parties;  // all parties (no filter)

  var now      = new Date();
  var curMonth = parseInt(month || (now.getMonth() + 1));
  var curYear  = parseInt(year  || now.getFullYear());

  // MTD orders (cancelled excluded)
  var monthOrders = orders.filter(function(o){
    if (o.status === 'Cancelled') return false;
    var d = parseDate(o.date);
    return d.getMonth()+1 === curMonth && d.getFullYear() === curYear;
  });

  // Today (IST)
  var today = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  var todayOrders = orders.filter(function(o){
    if (o.status === 'Cancelled') return false;
    var ds = String(o.date);
    if (/^\d{4}-\d{2}-\d{2}$/.test(ds)) return ds === today;
    var parsed = parseDate(ds);
    return Utilities.formatDate(parsed, 'Asia/Kolkata', 'yyyy-MM-dd') === today;
  });

  // YTD orders (cancelled excluded)
  var ytdOrders = orders.filter(function(o){
    if (o.status === 'Cancelled') return false;
    return parseDate(o.date).getFullYear() === curYear;
  });

  var mtdSales   = monthOrders.reduce(function(s,o){return s+(o.gt||0);},0);
  var ytdSales   = ytdOrders.reduce(function(s,o){return s+(o.gt||0);},0);
  var todaySales = todayOrders.reduce(function(s,o){return s+(o.gt||0);},0);
  var todayShops = countUnique(todayOrders.map(function(o){return o.bid;}));
  var mtdShops   = countUnique(monthOrders.map(function(o){return o.bid;}));

  // Product mix (MTD)
  var prodMap = {};
  monthOrders.forEach(function(o){
    o.items.forEach(function(it){
      prodMap[it.code] = (prodMap[it.code]||0) + it.total;
    });
  });
  var prodMix = [];
  for (var code in prodMap) prodMix.push({code:code, value:prodMap[code]});
  prodMix.sort(function(a,b){return b.value-a.value;});

  // Target
  var target = 50000;
  if (smId) {
    var sm = salesmen.filter(function(s){return s.id===smId;})[0];
    if (sm) target = sm.target || 50000;
  }

  // New parties THIS MONTH added BY this salesman
  var newPartiesList = allParties.filter(function(p){
    if (smId && p.addedBy !== smId) return false;
    if (!p.at) return false;
    var d = parseDate(p.at);
    return d.getMonth()+1 === curMonth && d.getFullYear() === curYear;
  });

  // Admin: salesman breakdown
  var smBreakdown = [];
  if (!smId) {
    var smMap = {};
    allOrdersList.filter(function(o){return o.status!=='Cancelled';}).forEach(function(o){
      if (!smMap[o.smId]) smMap[o.smId] = {smId:o.smId, smName:o.smName, orders:0, rev:0};
      smMap[o.smId].orders++;
      smMap[o.smId].rev += (o.gt||0);
    });
    for (var sid in smMap) smBreakdown.push(smMap[sid]);
  }

  // Get SM targets
  var smRec = null;
  if (smId) {
    var smList = getSalesmen().salesmen;
    for (var si=0; si<smList.length; si++) {
      if (smList[si].id === smId) { smRec = smList[si]; break; }
    }
  }
  var shopsTarget    = smRec ? (smRec.shopsTarget   ||0) : 0;
  var partiesTarget  = smRec ? (smRec.partiesTarget  ||0) : 0;
  var prodMixTargets = smRec ? (smRec.prodMixTargets ||{}) : {};
  // Daily targets — based on 26 working days per month
  var dailySalesTarget = smRec && smRec.dailyTarget > 0
    ? smRec.dailyTarget
    : (target > 0 ? Math.round(target / 26) : 0);
  var dailyShopsTarget = smRec && smRec.dailyShopsTarget > 0
    ? smRec.dailyShopsTarget
    : (shopsTarget > 0 ? Math.max(1, Math.round(shopsTarget / 26)) : 0);
  var visitTarget = smRec ? (smRec.visitTarget||50) : 50;

  return {
    ok: true,
    summary: {
      mtdSales        : mtdSales,
      ytdSales        : ytdSales,
      todaySales      : todaySales,
      todayShops      : todayShops,
      mtdShops        : mtdShops,
      target          : target,
      mtdPct          : Math.round((mtdSales / (target||1)) * 100),
      ytdPct          : Math.round((ytdSales / ((target||1) * 12)) * 100),
      newParties      : newPartiesList.length,
      totalRev        : ytdSales,
      totalOrders     : orders.length,
      totalComm       : orders.reduce(function(s,o){return s+(o.tc||0);},0),
      shopsTarget     : shopsTarget,
      partiesTarget   : partiesTarget,
      prodMixTargets  : prodMixTargets,
      dailySalesTarget: dailySalesTarget,
      dailyShopsTarget: dailyShopsTarget,
      visitTarget: visitTarget
    },
    productMix   : prodMix.slice(0, 7),
    recentOrders : orders.slice(0, 10),
    salesmen     : smBreakdown.length > 0 ? smBreakdown : salesmen,
    newParties   : newPartiesList   // full array for KPI modal
  };
}

// ════════════════════════════════════════════════════════════════════════
// UPDATE ORDER STATUS
// ════════════════════════════════════════════════════════════════════════

function updateStatus(orderId, status) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_ORDERS);
  if (!sheet) return {ok:false, error:'Sheet not found'};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  for (var i=0; i<rows.length; i++) {
    if (rows[i][0] === orderId) {
      sheet.getRange(i+2, 4).setValue(status);
    }
  }
  return {ok:true};
}

// ════════════════════════════════════════════════════════════════════════
// SAVE / UPDATE PRODUCT (admin master management)
// Creates a new row if code not found; updates existing row if found.
// ════════════════════════════════════════════════════════════════════════

function saveProductMaster(p) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_PRODUCTS, HDR_PRODUCTS);
  var commVal = p.comm || 0;
  var newRow = [p.code, p.name, p.category||'', p.price, commVal, p.active===false?'No':'Yes'];

  // If code was renamed, find old row by oldCode and update it
  var searchCode = (p.oldCode && p.oldCode !== p.code) ? p.oldCode : p.code;

  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0; i<rows.length; i++) {
      if (String(rows[i][0]).trim() === String(searchCode).trim()) {
        sheet.getRange(i+2, 1, 1, HDR_PRODUCTS.length).setValues([newRow]);
        return {ok:true, action:'updated', code:p.code};
      }
    }
  }
  // Not found — append as new
  sheet.appendRow(newRow);
  return {ok:true, action:'created', code:p.code};
}

// ════════════════════════════════════════════════════════════════════════
// PARTY ACCESS CONTROL
// Stores per-salesman access flags in a sheet called SM_Access.
// Row format: SM_ID | partyDetails (Yes/No)
// ════════════════════════════════════════════════════════════════════════

var TAB_ACCESS = 'SM_Access';
var HDR_ACCESS = ['SM_ID', 'Party Details'];

function getPartyAccess() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_ACCESS);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, access:{}};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_ACCESS.length).getValues();
  var access = {};
  for (var i=0; i<rows.length; i++) {
    var smId = String(rows[i][0]).trim();
    if (!smId) continue;
    access[smId] = rows[i][1] !== 'No';  // true = has access, false = hidden
  }
  return {ok:true, access:access};
}

function savePartyAccess(accessMap) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_ACCESS, HDR_ACCESS);

  // Read existing rows to update in-place
  var existing = {};
  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_ACCESS.length).getValues();
    for (var i=0; i<rows.length; i++) {
      var sid = String(rows[i][0]).trim();
      if (sid) existing[sid] = i+2;  // rowNum
    }
  }

  for (var smId in accessMap) {
    var val = accessMap[smId] === false ? 'No' : 'Yes';
    if (existing[smId]) {
      sheet.getRange(existing[smId], 2).setValue(val);
    } else {
      sheet.appendRow([smId, val]);
    }
  }
  return {ok:true};
}


// ════════════════════════════════════════════════════════════════════════
// BULK IMPORT PARTIES
// Imports multiple parties at once from admin route plan import.
// Skips duplicates by ID. Returns count of actually imported parties.
// ════════════════════════════════════════════════════════════════════════

function bulkImportParties(parties, smId) {
  if (!parties || !parties.length) return {ok:false, error:'No parties provided'};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_PARTIES, HDR_PARTIES);
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm');

  // Get existing IDs to skip duplicates
  var existing = {};
  if (sheet.getLastRow() >= 2) {
    var existRows = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (var i=0; i<existRows.length; i++) {
      existing[String(existRows[i][0]).trim()] = true;
    }
  }

  var imported = 0;
  var skipped = 0;
  var batchSize = 50;
  var batch = [];

  for (var j=0; j<parties.length; j++) {
    var p = parties[j];
    if (!p.id || existing[String(p.id).trim()]) { skipped++; continue; }
    batch.push([
      p.id, p.name, p.contact||'', p.phone||'', p.address||'',
      p.route||'', p.territory||'Ahmedabad', p.gst||'', p.terms||'',
      p.credit||0, p.outstanding||0, p.addedBy||smId||'', now,
      p.status||'Active', p.lat||'', p.lng||'', p.photo||'',
      p.whatsapp||'', p.email||'', p.notes||'', p.day||''
    ]);
    if (batch.length >= batchSize) {
      sheet.getRange(sheet.getLastRow()+1, 1, batch.length, HDR_PARTIES.length).setValues(batch);
      imported += batch.length;
      batch = [];
    }
  }
  if (batch.length > 0) {
    sheet.getRange(sheet.getLastRow()+1, 1, batch.length, HDR_PARTIES.length).setValues(batch);
    imported += batch.length;
  }

  return {ok:true, imported:imported, skipped:skipped};
}

// ════════════════════════════════════════════════════════════════════════
// SAVE PARTY ALLOCATIONS
// Sets AddedBy for all allocated party IDs to smId.
// Clears smId from any parties NOT in the list (unallocated).
// ════════════════════════════════════════════════════════════════════════

function savePartyAllocations(smId, partyIds) {
  if (!smId) return {ok:false, error:'smId required'};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_PARTIES);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, updated:0};

  var assignedSet = {};
  for (var i=0; i<partyIds.length; i++) assignedSet[String(partyIds[i]).trim()] = true;

  // Write SM allocation to col 12 (Added By) ONLY
  // Never touch col 22 (Dist Alloc)
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  // Read col 12 separately
  var col12 = sheet.getRange(2,12,sheet.getLastRow()-1,1).getValues();
  var updated = 0;
  for (var r=0; r<rows.length; r++) {
    var pid = String(rows[r][0]).trim();
    if (!pid) continue;
    var currentSm = String(col12[r][0]||'').trim();
    // Skip rows that have DIST: values in col 12 (legacy data) — don't overwrite
    if (currentSm.indexOf('DIST:') === 0) continue;
    if (assignedSet[pid] && currentSm !== smId) {
      sheet.getRange(r+2, 12).setValue(smId);
      updated++;
    } else if (!assignedSet[pid] && currentSm === smId) {
      sheet.getRange(r+2, 12).setValue('');
      updated++;
    }
  }
  return {ok:true, updated:updated};
}


function saveDistPartyAllocations(distId, partyIds, data) {
  data = data || {};
  if (!distId) return {ok:false, error:'distId required'};
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Use separate sheet — never touch New_Parties
  var sheet = ss.getSheetByName('Dist_Alloc');
  if (!sheet) {
    sheet = ss.insertSheet('Dist_Alloc');
    sheet.getRange(1,1,1,2).setValues([['Party ID','Dist ID']]);
    sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  var lastRow = sheet.getLastRow();
  var existRows = lastRow >= 2 ? sheet.getRange(2,1,lastRow-1,2).getValues() : [];

  if (data.singleParty) {
    // Single party mode — just update this one party's dist alloc
    var partyId = data.partyId;
    var found = false;
    for (var i = existRows.length-1; i >= 0; i--) {
      if (String(existRows[i][0]||'').trim() === partyId) {
        if (distId === '__CLEAR__' || !partyIds || !partyIds.length) {
          sheet.deleteRow(i+2); // remove allocation
        } else {
          sheet.getRange(i+2, 2).setValue(distId); // update
        }
        found = true; break;
      }
    }
    if (!found && partyIds && partyIds.length && distId !== '__CLEAR__') {
      sheet.appendRow([partyId, distId]); // add new
    }
    return {ok:true, mode:'single', partyId:partyId};
  }

  // Bulk mode — clear all rows for this distributor, rewrite
  for (var i = existRows.length-1; i >= 0; i--) {
    if (String(existRows[i][1]||'').trim() === distId) {
      sheet.deleteRow(i+2);
    }
  }
  if (partyIds && partyIds.length > 0) {
    var newRows = partyIds.map(function(pid){ return [pid, distId]; });
    sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, 2).setValues(newRows);
  }
  return {ok:true, distId:distId, count:partyIds?partyIds.length:0};
}

function getDistPartyAlloc(distId) {
  // Returns set of party IDs allocated to a distributor
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Dist_Alloc');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  var result = [];
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][1]||'').trim() === distId) {
      result.push(String(rows[i][0]||'').trim());
    }
  }
  return result;
}

function getAllDistAlloc() {
  // Returns map of partyId -> distId for all allocations
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Dist_Alloc');
  if (!sheet || sheet.getLastRow() < 2) return {};
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  var map = {};
  for (var i=0; i<rows.length; i++) {
    var pid = String(rows[i][0]||'').trim();
    var did = String(rows[i][1]||'').trim();
    if (pid && did) map[pid] = did;
  }
  return map;
}


// ════════════════════════════════════════════════════════════════════════
// VERIFY PIN — used by login screens
// ════════════════════════════════════════════════════════════════════════

function verifyPin(id, pin, role) {
  if (!id || !pin) return {ok:false, error:'ID and PIN required'};

  if (role === 'sm') {
    var sms = getSalesmen().salesmen;
    var sm = null;
    for (var i=0; i<sms.length; i++) { if (sms[i].id === id) { sm=sms[i]; break; } }
    if (!sm) return {ok:false, error:'Salesman ID not found'};
    if (sm.active === false) return {ok:false, error:'Account is deactivated'};
    if (!sm.pin) return {ok:false, error:'No PIN set — contact admin'};
    if (String(sm.pin).trim() !== String(pin).trim()) return {ok:false, error:'Incorrect PIN'};
    return {ok:true, id:sm.id, name:sm.name, distributor:sm.distributor, target:sm.target};
  }

  if (role === 'dist') {
    var dists = getDistributors().distributors;
    var dist = null;
    for (var i=0; i<dists.length; i++) { if (dists[i].id === id) { dist=dists[i]; break; } }
    if (!dist) return {ok:false, error:'Distributor ID not found'};
    if (dist.active === false) return {ok:false, error:'Account is deactivated'};
    if (!dist.pin) return {ok:false, error:'No PIN set — contact admin'};
    if (String(dist.pin).trim() !== String(pin).trim()) return {ok:false, error:'Incorrect PIN'};
    return {ok:true, id:dist.id, name:dist.name};
  }

  if (role === 'admin') {
    // Admin PIN stored in Script Properties for security
    var adminPin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '0000';
    if (String(pin).trim() !== adminPin) return {ok:false, error:'Incorrect admin PIN'};
    return {ok:true, id:'ADMIN', name:'Admin'};
  }

  return {ok:false, error:'Unknown role'};
}





// ════════════════════════════════════════════════════════════════════════
// VISIT LOG
// Sheet: Visit_Log
// Cols: Visit ID | Date | Time | SM ID | SM Name | Party ID | Party Name |
//       Outcome | Notes | Latitude | Longitude | Accuracy
// ════════════════════════════════════════════════════════════════════════

var TAB_VISITS = 'Visit_Log';
var HDR_VISITS = [
  'Visit ID','Date','Time','SM ID','SM Name',
  'Party ID','Party Name','Outcome','Notes',
  'Latitude','Longitude','Accuracy'
];

function saveVisit(v) {
  var ss   = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_VISITS, HDR_VISITS);
  var now  = new Date();
  var date = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
  var time = Utilities.formatDate(now, 'Asia/Kolkata', 'HH:mm:ss');
  var vid  = 'V-' + String(now.getTime());
  sheet.appendRow([
    vid, date, time,
    String(v.smId||'').trim(), String(v.smName||'').trim(),
    String(v.partyId||'').trim(), String(v.partyName||'').trim(),
    String(v.outcome||'').trim(), String(v.notes||'').trim(),
    v.lat||'', v.lng||'', v.accuracy||''
  ]);
  return {ok:true, visitId:vid};
}

function getVisits(smId, distId, dateFrom, dateTo) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_VISITS);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, visits:[]};

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_VISITS.length).getValues();
  var visits = [];
  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    if (smId   && String(r[3]).trim() !== String(smId).trim())   continue;
    var rDate = r[1] instanceof Date
      ? Utilities.formatDate(r[1], 'Asia/Kolkata', 'yyyy-MM-dd')
      : String(r[1]||'').trim();
    if (dateFrom && rDate < dateFrom) continue;
    if (dateTo   && rDate > dateTo)   continue;
    // Format date and time as plain strings — avoid Date object serialization bug
    var visitDate = r[1] instanceof Date
      ? Utilities.formatDate(r[1], 'Asia/Kolkata', 'yyyy-MM-dd')
      : String(r[1]||'').trim();
    var visitTime = r[2] instanceof Date
      ? Utilities.formatDate(r[2], 'Asia/Kolkata', 'HH:mm:ss')
      : String(r[2]||'').trim();
    visits.push({
      id:r[0], date:visitDate, time:visitTime,
      smId:String(r[3]||'').trim(), smName:String(r[4]||'').trim(),
      partyId:String(r[5]||'').trim(), partyName:String(r[6]||'').trim(),
      outcome:String(r[7]||'').trim(), notes:String(r[8]||'').trim(),
      lat:String(r[9]||'').trim(), lng:String(r[10]||'').trim(),
      accuracy:String(r[11]||'').trim()
    });
  }
  return {ok:true, visits:visits};
}

function getVisitStats(smId) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_VISITS);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, stats:{mtdVisits:0,ytdVisits:0,mtdOrders:0,ytdOrders:0}};

  var now = new Date();
  var curMonth = now.getMonth() + 1;
  var curYear  = now.getFullYear();
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_VISITS.length).getValues();

  var mtdVisits=0, ytdVisits=0, mtdOrders=0, ytdOrders=0;
  var mtdParties=new Set(), ytdParties=new Set();

  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    if (smId && String(r[3]).trim() !== String(smId).trim()) continue;
    var d = r[1] ? String(r[1]) : '';
    if (!d) continue;
    var parts = d.split('-');
    if (parts.length < 3) continue;
    var yr = parseInt(parts[0]), mo = parseInt(parts[1]);
    var isYtd = yr === curYear;
    var isMtd = isYtd && mo === curMonth;
    var outcome = String(r[7]||'').trim();
    var pid = String(r[5]||'').trim();
    if (isMtd){ mtdVisits++; mtdParties.add(pid); if(outcome==='Order Placed') mtdOrders++; }
    if (isYtd){ ytdVisits++; ytdParties.add(pid); if(outcome==='Order Placed') ytdOrders++; }
  }
  return {ok:true, stats:{
    mtdVisits:mtdVisits, ytdVisits:ytdVisits,
    mtdOrders:mtdOrders, ytdOrders:ytdOrders,
    mtdUniqueShops:mtdParties.size, ytdUniqueShops:ytdParties.size
  }};
}

// ════════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════════

function makeSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var r = sheet.getRange(1,1,1,headers.length);
    r.setValues([headers]);
    r.setBackground('#1F3864');
    r.setFontColor('#FFFFFF');
    r.setFontWeight('bold');
    sheet.setFrozenRows(1);
    for (var i=0; i<headers.length; i++) {
      sheet.setColumnWidth(i+1, 140);
    }
  }
  return sheet;
}

function colorRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  for (var r=2; r<=last; r++) {
    var color = (r%2===0) ? '#EBF3FB' : '#FFFFFF';
    sheet.getRange(r,1,1,sheet.getLastColumn()).setBackground(color);
  }
}

function parseDate(dateStr) {
  if (!dateStr) return new Date();
  if (Object.prototype.toString.call(dateStr) === '[object Date]') return dateStr;
  dateStr = dateStr.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    var p = dateStr.split('-');
    return new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
  }
  var parts = dateStr.split('-');
  if (parts.length === 3) {
    var months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,
                  Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
    if (months[parts[1]] !== undefined) {
      return new Date(parseInt(parts[2]), months[parts[1]], parseInt(parts[0]));
    }
  }
  return new Date(dateStr);
}

function countUnique(arr) {
  var u = {};
  for (var i=0; i<arr.length; i++) u[arr[i]] = true;
  var c = 0;
  for (var k in u) c++;
  return c;
}

// ════════════════════════════════════════════════════════════════════════
// UPDATE PARTY FULL (Admin direct edit)
// ════════════════════════════════════════════════════════════════════════

function updatePartyFull(p) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_PARTIES);
  if (!sheet) return {ok:false, error:'Parties sheet not found'};
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_PARTIES.length).getValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]).trim() !== String(p.id).trim()) continue;
    var rowNum = i + 2;
    sheet.getRange(rowNum, 1, 1, HDR_PARTIES.length).setValues([[
      p.id, p.name, p.contact||'', p.phone||'', p.address||'',
      p.route||'', p.territory||'', p.gst||'', p.terms||'',
      p.credit||0, p.outstanding||0,
      p.addedBy!==undefined ? p.addedBy : rows[i][11]||'',
      rows[i][12],
      p.status||'Active',
      p.lat||'', p.lng||'', p.photo||'',
      p.whatsapp||'', p.email||'', p.notes||'', p.day||''
    ]]);
    return {ok:true, id:p.id};
  }
  return {ok:false, error:'Party not found: '+p.id};
}

// ════════════════════════════════════════════════════════════════════════
// UPDATE PARTY STATUS (Hold / Active / Inactive)
// ════════════════════════════════════════════════════════════════════════

function updatePartyStatus(partyId, status) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_PARTIES);
  if (!sheet) return {ok:false, error:'Parties sheet not found'};
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_PARTIES.length).getValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]).trim() === String(partyId).trim()) {
      sheet.getRange(i+2, 14).setValue(status);
      return {ok:true, partyId:partyId, status:status};
    }
  }
  return {ok:false, error:'Party not found'};
}

// ════════════════════════════════════════════════════════════════════════
// DEACTIVATE / REACTIVATE SALESMAN
// ════════════════════════════════════════════════════════════════════════

function deactivateSalesman(smId, active) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_SALESMEN);
  if (!sheet) return {ok:false, error:'Salesmen sheet not found'};
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_SALESMEN.length).getValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]).trim() === String(smId).trim()) {
      sheet.getRange(i+2, 8).setValue(active ? 'Yes' : 'No');
      return {ok:true, smId:smId, active:active};
    }
  }
  return {ok:false, error:'Salesman not found'};
}

// ════════════════════════════════════════════════════════════════════════
// DEACTIVATE / REACTIVATE DISTRIBUTOR
// ════════════════════════════════════════════════════════════════════════

function deactivateDistributor(distId, active) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_DISTRIBUTORS);
  if (!sheet) return {ok:false, error:'Distributors sheet not found'};
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_DISTRIBUTORS.length).getValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]).trim() === String(distId).trim()) {
      sheet.getRange(i+2, 6).setValue(active ? 'Yes' : 'No');
      return {ok:true, distId:distId, active:active};
    }
  }
  return {ok:false, error:'Distributor not found'};
}

// ════════════════════════════════════════════════════════════════════════
// EDIT REQUESTS (Salesman suggests → Admin approve/reject)
// ════════════════════════════════════════════════════════════════════════

var TAB_EDIT_REQUESTS = 'Party_Edit_Requests';
var HDR_EDIT_REQUESTS = [
  'Request ID','Party ID','Party Name','SM ID',
  'Changes JSON','Note','Status','Submitted At','Resolved At'
];

function getEditRequests() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_EDIT_REQUESTS, HDR_EDIT_REQUESTS);
  if (sheet.getLastRow() < 2) return {ok:true, requests:[]};
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_EDIT_REQUESTS.length).getValues();
  var reqs = [];
  for (var i=0; i<rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var changes = {};
    try { changes = JSON.parse(r[4]||'{}'); } catch(e) {}
    reqs.push({
      id:String(r[0]).trim(), partyId:String(r[1]).trim(),
      partyName:String(r[2]).trim(), smId:String(r[3]).trim(),
      changes:changes, note:String(r[5]||'').trim(),
      status:String(r[6]||'Pending').trim(),
      at:String(r[7]||'').trim(), resolvedAt:String(r[8]||'').trim()
    });
  }
  return {ok:true, requests:reqs};
}

function submitEditRequest(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = makeSheet(ss, TAB_EDIT_REQUESTS, HDR_EDIT_REQUESTS);
  var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm');
  var reqId = 'ER-'+new Date().getTime();
  sheet.appendRow([
    reqId, data.partyId||'', data.partyName||'', data.smId||'',
    JSON.stringify(data.changes||{}), data.note||'', 'Pending', now, ''
  ]);
  return {ok:true, reqId:reqId};
}

function resolveEditRequest(reqId, resolution) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_EDIT_REQUESTS);
  if (!sheet) return {ok:false, error:'Edit requests sheet not found'};
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_EDIT_REQUESTS.length).getValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]).trim() !== String(reqId).trim()) continue;
    var rowNum = i + 2;
    var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm');
    sheet.getRange(rowNum, 7).setValue(resolution);
    sheet.getRange(rowNum, 9).setValue(now);
    if (resolution === 'Approved') {
      var changes = {};
      try { changes = JSON.parse(rows[i][4]||'{}'); } catch(e) {}
      var partyId = String(rows[i][1]).trim();
      var partySheet = ss.getSheetByName(TAB_PARTIES);
      if (partySheet && partySheet.getLastRow() >= 2) {
        var pRows = partySheet.getRange(2,1,partySheet.getLastRow()-1,HDR_PARTIES.length).getValues();
        for (var j=0; j<pRows.length; j++) {
          if (String(pRows[j][0]).trim() !== partyId) continue;
          var pRow = j + 2;
          var colMap = {name:2,contact:3,phone:4,address:5,route:6,territory:7,gst:8,terms:9,notes:20,whatsapp:18,email:19};
          Object.keys(changes).forEach(function(k) {
            if (colMap[k] && changes[k].new !== undefined)
              partySheet.getRange(pRow, colMap[k]).setValue(changes[k].new);
            if (k==='gps' && changes[k].lat) {
              partySheet.getRange(pRow, 15).setValue(changes[k].lat);
              partySheet.getRange(pRow, 16).setValue(changes[k].lng);
            }
          });
          break;
        }
      }
    }
    return {ok:true, reqId:reqId, resolution:resolution};
  }
  return {ok:false, error:'Request not found'};
}
