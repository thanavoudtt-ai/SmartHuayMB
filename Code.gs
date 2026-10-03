const CONFIG = {
  SPREADSHEET_ID: '1dxzettk07Qmn9bGP6XeRfpxGAMf5NrvEQ3MdG-_TvDE',
  SHEET_NAME: 'DATA SMART HUAY',
  TZ: 'Asia/Vientiane'
};

const HEADERS = [
  'Bill ID','Created At','Updated At','Customer','Lottery','Digits','Number',
  'Position','Amount','Payment Method','Debt Status','Paid At','Paid Method','Status',
  'Agent ID','Agent Name','Seller Percent'
];
const AGENT_SHEET = 'SMART HUAY SETTINGS';

function doGet() {
  ensureSheet_();
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('SMART HUAY')
    .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureSheet_(){
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if(!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);
  if(sh.getLastRow() === 0){
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),HEADERS.length)).getValues()[0];
    HEADERS.forEach((h,i)=>{ if(current[i] !== h) sh.getRange(1,i+1).setValue(h); });
  }
  ensureAgentSheet_();
  return sh;
}

function ensureAgentSheet_(){
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(AGENT_SHEET);
  if(!sh) sh = ss.insertSheet(AGENT_SHEET);
  const headers=['Agent ID','Agent Name','Seller Percent','Active','Created At','Updated At'];
  if(sh.getLastRow()===0){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getAgents(includeInactive){
  const sh=ensureAgentSheet_(), v=sh.getDataRange().getValues();
  if(v.length<2) return [];
  return v.slice(1).filter(r=>r[0] && (includeInactive || r[3]===true || String(r[3]).toUpperCase()==='TRUE'))
    .map(r=>({id:String(r[0]),name:String(r[1]),percent:Number(r[2]||0),active:r[3]===true||String(r[3]).toUpperCase()==='TRUE'}));
}

function saveAgent(a){
  const sh=ensureAgentSheet_(), now=new Date(), name=String(a.name||'').trim(), pct=Number(a.percent);
  if(!name) throw new Error('ກະລຸນາໃສ່ຊື່ແມ່ຫວຍ');
  if(!isFinite(pct)||pct<0||pct>100) throw new Error('ເປີເຊັນຕ້ອງຢູ່ 0-100');
  if(a.id){
    const v=sh.getDataRange().getValues();
    for(let i=1;i<v.length;i++) if(String(v[i][0])===String(a.id)){
      sh.getRange(i+1,2,1,5).setValues([[name,pct,a.active!==false,v[i][4]||now,now]]);
      return {ok:true,id:String(a.id)};
    }
  }
  const id='AG-'+Utilities.formatDate(now,CONFIG.TZ,'yyyyMMddHHmmss');
  sh.appendRow([id,name,pct,a.active!==false,now,now]);
  return {ok:true,id};
}

function toggleAgent(id,active){
  const sh=ensureAgentSheet_(),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++) if(String(v[i][0])===String(id)){
    sh.getRange(i+1,4).setValue(!!active); sh.getRange(i+1,6).setValue(new Date()); return {ok:true};
  }
  throw new Error('ບໍ່ພົບແມ່ຫວຍ');
}

function deleteAgent(id){
  const sh=ensureAgentSheet_(),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++) if(String(v[i][0])===String(id)){ sh.deleteRow(i+1); return {ok:true}; }
  throw new Error('ບໍ່ພົບແມ່ຫວຍ');
}

function nextBillId_(){
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const props = PropertiesService.getScriptProperties();
    const day = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd');
    const key = 'BILL_SEQ_' + day;
    const n = Number(props.getProperty(key) || 0) + 1;
    props.setProperty(key, String(n));
    return `SH-${day}-${String(n).padStart(3,'0')}`;
  } finally { lock.releaseLock(); }
}

function validateBillPayload_(payload){
  if(!payload || !payload.customer || !Array.isArray(payload.items) || !payload.items.length) throw new Error('ຂໍ້ມູນບິນບໍ່ຄົບ');
  payload.items.forEach(it=>{ if(!it.number || Number(it.amount||0)<=0) throw new Error('ລາຍການບິນບໍ່ຖືກຕ້ອງ'); });
}
function receiptStamp_(d){ return Utilities.formatDate(d||new Date(),CONFIG.TZ,'dd/MM/yyyy HH:mm'); }
function saveBill(payload){
  validateBillPayload_(payload);
  const billId=payload.billId||nextBillId_();
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sh=ensureSheet_(), now=new Date();
    const rows=payload.items.map(it=>[billId,now,now,String(payload.customer).trim(),it.lottery,it.digits,it.number,it.position||'-',Number(it.amount||0),payload.paymentMethod||'ເງິນສົດ',payload.debt?'ຕິດໜີ້':'ຊຳລະແລ້ວ','','','ACTIVE',payload.agentId||'',payload.agentName||'',Number(payload.sellerPercent||0)]);
    sh.getRange(sh.getLastRow()+1,1,rows.length,HEADERS.length).setValues(rows); SpreadsheetApp.flush();
    return {ok:true,billId,createdAt:receiptStamp_(now),receiptAt:receiptStamp_(now)};
  } finally { lock.releaseLock(); }
}

function getBills(dateStr){
  const sh = ensureSheet_();
  const values = sh.getDataRange().getValues();
  if(values.length < 2) return [];
  const target = dateStr || Utilities.formatDate(new Date(),CONFIG.TZ,'yyyy-MM-dd');
  const map = {};
  values.slice(1).forEach((r,i)=>{
    if(r[13] === 'DELETED') return;
    const d = r[1] instanceof Date ? Utilities.formatDate(r[1],CONFIG.TZ,'yyyy-MM-dd') : '';
    if(d !== target) return;
    const id = r[0];
    if(!map[id]) map[id] = {billId:id,createdAt:r[1],customer:r[3],paymentMethod:r[9],debtStatus:r[10],agentId:r[14]||'',agentName:r[15]||'',sellerPercent:Number(r[16]||0),items:[],rowNumbers:[]};
    map[id].items.push({lottery:r[4],digits:r[5],number:r[6],position:r[7],amount:Number(r[8]||0)});
    map[id].rowNumbers.push(i+2);
  });
  return Object.values(map).map(b=>{
    b.total=b.items.reduce((s,x)=>s+x.amount,0);
    b.createdAt = b.createdAt instanceof Date ? Utilities.formatDate(b.createdAt,CONFIG.TZ,'dd/MM/yyyy HH:mm') : String(b.createdAt);
    return b;
  }).sort((a,b)=>b.billId.localeCompare(a.billId));
}

function getDebtBills(){
  const sh = ensureSheet_();
  const values = sh.getDataRange().getValues();
  const map={};
  values.slice(1).forEach((r,i)=>{
    if(r[13] === 'DELETED' || r[10] !== 'ຕິດໜີ້') return;
    const id=r[0];
    if(!map[id]) map[id]={billId:id,createdAt:r[1],customer:r[3],agentId:r[14]||'',agentName:r[15]||'',sellerPercent:Number(r[16]||0),items:[],rowNumbers:[]};
    map[id].items.push({lottery:r[4],digits:r[5],number:r[6],position:r[7],amount:Number(r[8]||0)});
    map[id].rowNumbers.push(i+2);
  });
  return Object.values(map).map(b=>{
    b.total=b.items.reduce((s,x)=>s+x.amount,0);
    b.createdAt=b.createdAt instanceof Date?Utilities.formatDate(b.createdAt,CONFIG.TZ,'dd/MM/yyyy HH:mm'):String(b.createdAt);
    return b;
  }).sort((a,b)=>b.billId.localeCompare(a.billId));
}

function updateBill(payload){
  validateBillPayload_(payload);
  if(!payload.billId) throw new Error('ບໍ່ພົບ Bill ID');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sh=ensureSheet_(), values=sh.getDataRange().getValues(), now=new Date(), old=[];
    for(let i=1;i<values.length;i++) if(String(values[i][0])===String(payload.billId) && values[i][13]!=='DELETED') old.push(i+1);
    if(!old.length) throw new Error('ບໍ່ພົບບິນ');
    const first=values[old[0]-1], created=first[1] instanceof Date?first[1]:now, payment=first[9]||payload.paymentMethod||'ເງິນສົດ', debt=first[10]||'ຊຳລະແລ້ວ', paidAt=first[11]||'', paidMethod=first[12]||'';
    const newRows=payload.items.map(it=>[payload.billId,created,now,String(payload.customer).trim(),it.lottery,it.digits,it.number,it.position||'-',Number(it.amount||0),payment,debt,paidAt,paidMethod,'ACTIVE',payload.agentId||'',payload.agentName||'',Number(payload.sellerPercent||0)]);
    const newStart=sh.getLastRow()+1; sh.getRange(newStart,1,newRows.length,HEADERS.length).setValues(newRows); SpreadsheetApp.flush();
    try{ old.forEach(r=>{sh.getRange(r,3).setValue(now);sh.getRange(r,14).setValue('DELETED')}); SpreadsheetApp.flush(); }
    catch(err){ sh.getRange(newStart,14,newRows.length,1).setValue('DELETED'); throw err; }
    return {ok:true,billId:payload.billId,createdAt:receiptStamp_(created),receiptAt:receiptStamp_(now)};
  } finally { lock.releaseLock(); }
}

function deleteBill(billId){
  const sh=ensureSheet_(), values=sh.getDataRange().getValues(), now=new Date();
  let n=0;
  for(let i=1;i<values.length;i++){
    if(values[i][0]===billId && values[i][13]!=='DELETED'){
      sh.getRange(i+1,3).setValue(now);
      sh.getRange(i+1,14).setValue('DELETED');
      n++;
    }
  }
  return {ok:true,rows:n};
}


function getSalesSummary(startDate,endDate){
  const sh=ensureSheet_(), v=sh.getDataRange().getValues();
  const start=startDate||Utilities.formatDate(new Date(),CONFIG.TZ,'yyyy-MM-dd');
  const end=endDate||start;
  const bills={}, byAgent={}, byDay={};
  v.slice(1).forEach(r=>{
    if(r[13]==='DELETED' || !r[0]) return;
    const d=r[1] instanceof Date?Utilities.formatDate(r[1],CONFIG.TZ,'yyyy-MM-dd'):'';
    if(d<start||d>end) return;
    const id=String(r[0]);
    if(!bills[id]) bills[id]={agentId:r[14]||'',agentName:r[15]||'ບໍ່ໄດ້ເລືອກ',pct:Number(r[16]||0),total:0,date:d};
    bills[id].total+=Number(r[8]||0);
  });
  Object.values(bills).forEach(b=>{
    const key=b.agentId||('legacy:'+b.agentName);
    if(!byAgent[key]) byAgent[key]={agentId:b.agentId,agentName:b.agentName,percent:b.pct,sales:0,bills:0};
    byAgent[key].sales+=b.total; byAgent[key].bills++;
    if(!byDay[b.date]) byDay[b.date]={date:b.date,sales:0,bills:0};
    byDay[b.date].sales+=b.total; byDay[b.date].bills++;
  });
  const rows=Object.values(byAgent).map(x=>{
    x.sellerShare=Math.round(x.sales*x.percent/100);
    x.agentPay=x.sales-x.sellerShare;
    return x;
  }).sort((a,b)=>b.sales-a.sales);
  return {
    rows,
    daily:Object.values(byDay).sort((a,b)=>a.date.localeCompare(b.date)),
    totals:rows.reduce((o,x)=>({sales:o.sales+x.sales,agentPay:o.agentPay+x.agentPay,sellerShare:o.sellerShare+x.sellerShare,bills:o.bills+x.bills}),{sales:0,agentPay:0,sellerShare:0,bills:0})
  };
}

function markDebt(billId){
  const sh=ensureSheet_(), values=sh.getDataRange().getValues(), now=new Date();
  let n=0;
  for(let i=1;i<values.length;i++){
    if(values[i][0]===billId && values[i][13]!=='DELETED'){
      sh.getRange(i+1,11).setValue('ຕິດໜີ້');
      sh.getRange(i+1,3).setValue(now);
      n++;
    }
  }
  return {ok:true,rows:n};
}

function settleDebt(billId, method){
  const sh=ensureSheet_(), values=sh.getDataRange().getValues(), now=new Date();
  let n=0;
  for(let i=1;i<values.length;i++){
    if(values[i][0]===billId && values[i][13]!=='DELETED'){
      sh.getRange(i+1,11).setValue('ຊຳລະແລ້ວ');
      sh.getRange(i+1,12).setValue(now);
      sh.getRange(i+1,13).setValue(method);
      sh.getRange(i+1,3).setValue(now);
      n++;
    }
  }
  return {ok:true,rows:n};
}
