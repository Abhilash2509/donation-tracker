// ================================================================
// STUDENT GROCERY DONATION TRACKER — Google Apps Script v6
// Author: abhilashnr90@gmail.com
// ================================================================
// SETUP: Run → runSetup  (once, grants Gmail + creates triggers)
// Then: Deploy → New Deployment → Web App (Execute as Me, Anyone)
// ================================================================

const ADMIN_EMAIL        = 'abhilashnr90@gmail.com';
const LOW_STOCK_THRESHOLD = 5;
const SURPLUS_THRESHOLD   = 10; // hide items from donor form at this balance

const SHEETS = {
  DONORS:          'Donor Master List',
  DONATIONS:       'Donation Records',
  RECIPIENTS:      'Recipient Master List',
  REQUESTS:        'Recipient Requests',
  HISTORY:         'History Log',
  ITEMS:           'Item Master',
  CUSTOM_REQUESTS: 'Custom Item Requests',
  NOTICES:         'Notice Board',
  WAITLIST:        'Waitlist',
};

// ════════════════════════════════════════════════
// ONE-TIME SETUP
// ════════════════════════════════════════════════
function runSetup() {
  sendEmail(ADMIN_EMAIL,
    '✅ Grocery Tracker — Setup complete',
    'Your Grocery Donation Tracker is configured.\n\nEmail notifications: ACTIVE\nAdmin: ' + ADMIN_EMAIL + '\nSetup time: ' + new Date().toString() + '\n\nAll systems ready.');

  // Clear and recreate triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['weeklyReportTrigger','dailyAdminReminderTrigger','weeklyDonorReportTrigger','weeklyLowStockTrigger'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Daily admin reminder at 10am PST (UTC-8, so 18:00 UTC)
  ScriptApp.newTrigger('dailyAdminReminderTrigger').timeBased().atHour(18).everyDays(1).create();

  // Weekly report to admin — Monday 8am PST (16:00 UTC)
  ScriptApp.newTrigger('weeklyReportTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(16).create();

  // Weekly donor report — Friday 9am PST (17:00 UTC)
  ScriptApp.newTrigger('weeklyDonorReportTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).create();

  // Weekly low stock alert — Monday 10am PST (18:00 UTC)
  ScriptApp.newTrigger('weeklyLowStockTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(18).create();

  Logger.log('Setup complete');
  return 'Setup complete. Check inbox: ' + ADMIN_EMAIL;
}

function testEmail() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const emails = getAllDonorEmails(ss);
  sendEmail(ADMIN_EMAIL, '🧪 Email test — Grocery Tracker',
    'Email working.\nDonor emails found: ' + emails.length + '\n' + emails.join('\n') + '\n\nTime: ' + new Date());
}

// ════════════════════════════════════════════════
// HTTP ROUTING
// ════════════════════════════════════════════════
function corsResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    const a = p.action || 'getData';
    let result;
    if      (a==='getData')               result=getAllData();
    else if (a==='addDonation')           result=addDonation(p);
    else if (a==='addRequest')            result=addRequest(p);
    else if (a==='addDonor')              result=addDonor(p);
    else if (a==='addRecipient')          result=addRecipient(p);
    else if (a==='updateDonationStatus')  result=updateDonationStatus(p);
    else if (a==='updateRequestStatus')   result=updateRequestStatus(p);
    else if (a==='addItem')               result=addItem(p);
    else if (a==='deleteItem')            result=deleteItem(p);
    else if (a==='addCustomRequest')      result=addCustomRequest(p);
    else if (a==='updateCustomRequest')   result=updateCustomRequest(p);
    else if (a==='submitFeedback')        result=submitFeedback(p);
    else if (a==='addNotice')             result=addNotice(p);
    else if (a==='addWaitlist')           result=addWaitlist(p);
    else if (a==='updateWaitlistStatus')  result=updateWaitlistStatus(p);
    else if (a==='deleteNotice')          result=deleteNotice(p);
    else if (a==='sendWeeklyReport')      result=sendWeeklyReport();
    else result={error:'Unknown action: '+a};
    return corsResponse(result);
  } catch(err) {
    Logger.log('doGet error: '+err.message+'\n'+err.stack);
    return corsResponse({error:err.message});
  }
}

// ════════════════════════════════════════════════
// READ
// ════════════════════════════════════════════════
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function sheetRows(name) {
    const ws = ss.getSheetByName(name);
    if (!ws) return [];
    const data = ws.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0].map(h=>String(h).trim());
    return data.slice(1).map((row,ri)=>{
      const obj={_row:ri+2};
      headers.forEach((h,i)=>{
        let v=row[i];
        if(v instanceof Date) v=Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');
        obj[h]=(v===null||v===undefined||v==='')?'':v;
      });
      return obj;
    }).filter(r=>Object.keys(r).filter(k=>k!=='_row').some(k=>r[k]!==''));
  }

  const donations=sheetRows(SHEETS.DONATIONS);
  const requests =sheetRows(SHEETS.REQUESTS);

  const invMap={};
  donations.forEach(r=>{
    if(String(r['Status']||'').toLowerCase()==='collected'){
      const item=String(r['Item']||'').trim();if(!item)return;
      if(!invMap[item])invMap[item]={donated:0,distributed:0};
      invMap[item].donated+=Number(r['Quantity'])||0;
    }
  });
  requests.forEach(r=>{
    if(String(r['Status']||'').toLowerCase()==='delivered'){
      const item=String(r['Item Requested']||'').trim();if(!item)return;
      if(!invMap[item])invMap[item]={donated:0,distributed:0};
      invMap[item].distributed+=Number(r['Qty Requested'])||0;
    }
  });
  const inventory=Object.entries(invMap).map(([item,v])=>({item,donated:v.donated,distributed:v.distributed,balance:v.donated-v.distributed}));

  return {
    donors:         sheetRows(SHEETS.DONORS),
    donations,
    recipients:     sheetRows(SHEETS.RECIPIENTS),
    requests,
    inventory,
    history:        sheetRows(SHEETS.HISTORY),
    items:          sheetRows(SHEETS.ITEMS).filter(r=>String(r['Active']||'').toLowerCase()!=='no'),
    customRequests: sheetRows(SHEETS.CUSTOM_REQUESTS),
    notices:        sheetRows(SHEETS.NOTICES),
    waitlist:       sheetRows(SHEETS.WAITLIST),
    lastUpdated:    new Date().toISOString(),
  };
}

// ════════════════════════════════════════════════
// DONATIONS
// ════════════════════════════════════════════════
function addDonation(body) {
  if(!body.item)     throw new Error('Missing required field: item');
  if(!body.quantity) throw new Error('Missing required field: quantity');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const donorId=String(body.donorId||'').trim();
  if(!donorId) throw new Error('Please provide your Donor ID.');
  verifyId(ss,SHEETS.DONORS,'Donor ID',donorId);

  const ws=getOrCreateSheet(ss,SHEETS.DONATIONS,['Donation ID','Date','Donor ID','Item','Quantity','Status','Submitted On','Collected On','Rejection Comment','Notes']);
  const donationId=generateTxnId(donorId,'D',ss,SHEETS.DONATIONS,'Donation ID');
  const now=new Date();
  ws.appendRow([donationId,fmt(now,'yyyy-MM-dd'),donorId,String(body.item),Number(body.quantity),'Submitted',fmt(now,'yyyy-MM-dd HH:mm'),'','',String(body.notes||'')]);
  logHistory(ss,'DONATION_SUBMITTED',donorId,body.item,body.quantity,now,'Submitted',donationId);

  // Acknowledgment to donor
  const donorEmail=getPersonEmail(ss,SHEETS.DONORS,'Donor ID',donorId);
  const donorName=getPersonName(ss,SHEETS.DONORS,'Donor ID',donorId);
  if(donorEmail) sendEmail(donorEmail,
    '🎁 Donation received — '+donationId,
    'Hi '+donorName+',\n\nThank you! Your donation has been recorded.\n\nDonation ID: '+donationId+'\nItem: '+body.item+'\nQuantity: '+body.quantity+'\nDate: '+fmt(now,'dd MMM yyyy')+'\n\nWe will contact you shortly to arrange collection from your address. You can use your Donation ID to track the status in the Donor Portal.\n\nThank you for your generosity!\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL);

  return {success:true,message:'Donation submitted — awaiting collection',donorId,donationId};
}

function updateDonationStatus(body) {
  if(!body.rowIndex) throw new Error('Missing rowIndex');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=ss.getSheetByName(SHEETS.DONATIONS);
  if(!ws) throw new Error('Donation Records sheet not found');
  const allData=ws.getDataRange().getValues();
  const headers=allData[0].map(h=>String(h).trim());
  const rowNum=Number(body.rowIndex);
  const status=String(body.status||'');
  const comment=String(body.comment||'');
  const rowArr=allData[rowNum-1];
  const rowObj={};headers.forEach((h,i)=>rowObj[h]=rowArr[i]);

  const setCol=(col,val)=>{const ci=headers.indexOf(col);if(ci>-1)ws.getRange(rowNum,ci+1).setValue(val);};
  setCol('Status',status);
  if(status==='Collected') setCol('Collected On',fmt(new Date(),'yyyy-MM-dd HH:mm'));
  if(status==='Rejected')  setCol('Rejection Comment',comment);

  logHistory(ss,'DONATION_'+status.toUpperCase().replace(/ /g,'_'),String(rowObj['Donor ID']||''),String(rowObj['Item']||''),rowObj['Quantity']||0,new Date(),status,String(rowObj['Donation ID']||''));
  emailDonationStatus(ss,rowObj,status,comment);

  return {success:true,message:'Donation marked as '+status,data:getAllData()};
}

// ════════════════════════════════════════════════
// REQUESTS
// ════════════════════════════════════════════════
function addRequest(body) {
  if(!body.itemRequested) throw new Error('Missing required field: itemRequested');
  if(!body.qtyRequested)  throw new Error('Missing required field: qtyRequested');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const recipientId=String(body.recipientId||'').trim();
  if(!recipientId) throw new Error('Please provide your Recipient ID.');
  verifyId(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
  const itemName=String(body.itemRequested).trim();

  // ── EFFECTIVE BALANCE CHECK ──
  // Balance = collected donations - delivered requests
  // Effective = balance - qty of active (pending/approved/delivery pending) requests
  const donWs=ss.getSheetByName(SHEETS.DONATIONS);
  const reqWs=ss.getSheetByName(SHEETS.REQUESTS);
  let donated=0,distributed=0,reserved=0;
  if(donWs){
    const ddata=donWs.getDataRange().getValues();
    const dh=ddata[0].map(h=>String(h).trim());
    const ic=dh.indexOf('Item'),sc=dh.indexOf('Status'),qc=dh.indexOf('Quantity');
    ddata.slice(1).forEach(r=>{if(String(r[ic]||'').trim()===itemName&&String(r[sc]||'').toLowerCase()==='collected')donated+=Number(r[qc])||0;});
  }
  if(reqWs){
    const rdata=reqWs.getDataRange().getValues();
    const rh=rdata[0].map(h=>String(h).trim());
    const ic=rh.indexOf('Item Requested'),sc=rh.indexOf('Status'),qc=rh.indexOf('Qty Requested'),ridc=rh.indexOf('Recipient ID');
    rdata.slice(1).forEach(r=>{
      const s=String(r[sc]||'').toLowerCase();
      if(String(r[ic]||'').trim()===itemName){
        if(s==='delivered') distributed+=Number(r[qc])||0;
        if(['pending','approved','delivery pending'].includes(s)) reserved+=Number(r[qc])||0;
      }
    });
  }
  const balance=donated-distributed;
  const effective=balance-reserved;

  // If nothing is available (balance=0 or all reserved), offer waitlist
  if(effective<=0&&balance<=0||effective<=0&&balance>0){
    return {
      success:false,
      waitlistNeeded:true,
      item:itemName,
      message:'RESERVED:'+itemName
    };
  }

  const ws=getOrCreateSheet(ss,SHEETS.REQUESTS,['Request ID','Request Date','Recipient ID','Item Requested','Qty Requested','Status','Pending On','Approved On','Delivery Pending On','Delivered On','Rejected On','Rejection Comment','ETA','Notes','Feedback']);
  const requestId=generateTxnId(recipientId,'R',ss,SHEETS.REQUESTS,'Request ID');
  const now=new Date();
  ws.appendRow([requestId,fmt(now,'yyyy-MM-dd'),recipientId,itemName,Number(body.qtyRequested),'Pending',fmt(now,'yyyy-MM-dd HH:mm'),'','','','','','',String(body.notes||''),'']);
  logHistory(ss,'REQUEST_SUBMITTED',recipientId,itemName,body.qtyRequested,now,'Pending',requestId);

  const recEmail=getPersonEmail(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
  const recName=getPersonName(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
  if(recEmail) sendEmail(recEmail,
    '🤲 Request received — '+requestId,
    'Hi '+recName+',\n\nYour request has been recorded and is pending admin review.\n\nRequest ID: '+requestId+'\nItem: '+itemName+'\nDate: '+fmt(now,'dd MMM yyyy')+'\nDelivery preference: '+(body.notes||'—')+'\n\nYou can track your request status in the Recipient Portal using your Request ID. You will receive an email at each status update.\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL);

  sendEmail(ADMIN_EMAIL,
    '⚡ Action needed — New request for '+itemName,
    'A new recipient request requires your review.\n\nRequest ID: '+requestId+'\nRecipient ID: '+recipientId+'\nItem: '+itemName+'\nStock: '+balance+' · Reserved: '+reserved+' · Effective available: '+effective+'\nDelivery preference: '+(body.notes||'—')+'\nDate: '+fmt(now,'dd MMM yyyy HH:mm')+'\n\nPlease log into the admin portal to approve or action this request.\n\n'+ADMIN_EMAIL);

  emailAllDonors_NewRequest(ss,itemName,recName,requestId);
  return {success:true,message:'Request submitted — pending admin review',recipientId,requestId};
}

function updateRequestStatus(body) {
  if(!body.rowIndex) throw new Error('Missing rowIndex');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=ss.getSheetByName(SHEETS.REQUESTS);
  if(!ws) throw new Error('Recipient Requests sheet not found');
  const allData=ws.getDataRange().getValues();
  const headers=allData[0].map(h=>String(h).trim());
  const rowNum=Number(body.rowIndex);
  const status=String(body.status||'');
  const comment=String(body.comment||'');
  const eta=String(body.eta||'');
  const rowArr=allData[rowNum-1];
  const rowObj={};headers.forEach((h,i)=>rowObj[h]=rowArr[i]);

  const setCol=(col,val)=>{const ci=headers.indexOf(col);if(ci>-1)ws.getRange(rowNum,ci+1).setValue(val);};
  setCol('Status',status);
  const ts=fmt(new Date(),'yyyy-MM-dd HH:mm');
  if(status==='Approved')        {setCol('Approved On',ts);if(eta)setCol('ETA',eta);}
  if(status==='Delivered')        setCol('Delivered On',ts);
  if(status==='Rejected')        {setCol('Rejected On',ts);setCol('Rejection Comment',comment);}

  logHistory(ss,'REQUEST_'+status.toUpperCase().replace(/ /g,'_'),String(rowObj['Recipient ID']||''),String(rowObj['Item Requested']||''),rowObj['Qty Requested']||0,new Date(),status,String(rowObj['Request ID']||''));
  emailRequestStatus(ss,rowObj,status,comment,eta);

  // When a request is Delivered, find the specific donation that funded it
  // and update waitlist
  if(status==='Delivered'||status==='Rejected'){
    const donationTxnId=findFundingDonation(ss,String(rowObj['Item Requested']||''));
    notifyWaitlist(ss,String(rowObj['Item Requested']||''),status==='Delivered'?donationTxnId:'');
  }

  return {success:true,message:'Request marked as '+status,data:getAllData()};
}

// ════════════════════════════════════════════════
// FEEDBACK
// ════════════════════════════════════════════════
function submitFeedback(body) {
  if(!body.requestId) throw new Error('Missing requestId');
  if(!body.feedback)  throw new Error('Missing feedback');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=ss.getSheetByName(SHEETS.REQUESTS);
  if(!ws) throw new Error('Requests sheet not found');

  let allData=ws.getDataRange().getValues();
  let headers=allData[0].map(h=>String(h).trim());

  // Auto-add Feedback column if missing (handles old sheets)
  let fbCol=headers.indexOf('Feedback');
  if(fbCol<0){
    const lastCol=ws.getLastColumn();
    ws.getRange(1,lastCol+1).setValue('Feedback');
    fbCol=lastCol; // 0-indexed = lastCol (was 1-indexed lastCol+1)
    // Refresh data after adding column
    allData=ws.getDataRange().getValues();
    headers=allData[0].map(h=>String(h).trim());
    fbCol=headers.indexOf('Feedback');
  }

  const idCol=headers.indexOf('Request ID');
  const notesCol=headers.indexOf('Notes'); // we need this to ensure we don't read Notes as Feedback

  for(let i=1;i<allData.length;i++){
    if(String(allData[i][idCol]).trim()===String(body.requestId).trim()){
      const feedbackText=String(body.feedback).trim();
      ws.getRange(i+1,fbCol+1).setValue(feedbackText);

      const recipientId=String(allData[i][headers.indexOf('Recipient ID')]||'');
      const item=String(allData[i][headers.indexOf('Item Requested')]||'');
      const deliveredDate=String(allData[i][headers.indexOf('Delivered On')]||'');
      const recName=getPersonName(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);

      // Notify donor(s) who donated this item with the feedback
      notifyDonorOfFeedback(ss,item,recName,feedbackText,body.requestId);

      // Also send confirmation email to recipient
      const recEmail=getPersonEmail(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
      if(recEmail){
        sendEmail(recEmail,
          '💬 Feedback received — thank you!',
          'Hi '+recName+',

Thank you for sharing your feedback!

Your feedback has been recorded and shared with the donor.

Your feedback: "'+feedbackText+'"

Your contribution to this community means a great deal.

Grocery Donation Tracker
'+ADMIN_EMAIL);
      }

      Logger.log('Feedback submitted for '+body.requestId+': '+feedbackText);
      return {success:true,message:'Feedback submitted. Thank you!'};
    }
  }
  throw new Error('Request ID not found: '+body.requestId);
}

function notifyDonorOfFeedback(ss,item,recipientName,feedback,requestId){
  // Notify all donors who donated this item (collected)
  const donWs=ss.getSheetByName(SHEETS.DONATIONS);
  if(!donWs)return;
  const donData=donWs.getDataRange().getValues();
  const donH=donData[0].map(h=>String(h).trim());
  const itemCol=donH.indexOf('Item');
  const statusCol=donH.indexOf('Status');
  const donorIdCol=donH.indexOf('Donor ID');
  const notified=new Set();
  donData.slice(1).forEach(row=>{
    if(String(row[itemCol]).trim().toLowerCase()===item.toLowerCase()&&String(row[statusCol]).trim().toLowerCase()==='collected'){
      const did=String(row[donorIdCol]).trim();
      if(did&&!notified.has(did)){
        notified.add(did);
        const email=getPersonEmail(ss,SHEETS.DONORS,'Donor ID',did);
        const name=getPersonName(ss,SHEETS.DONORS,'Donor ID',did);
        if(email) sendEmail(email,
          '💚 Your donation of '+item+' has made a difference!',
          'Hi '+name+',\n\nWonderful news! The '+item+' you donated was just delivered to a community member in need.\n\nHere is what they said:\n\n"'+feedback+'"\n\n— '+recipientName+'\n\nYour generosity truly makes a difference. Thank you so much!\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL);
      }
    }
  });
}

// ════════════════════════════════════════════════
// PEOPLE
// ════════════════════════════════════════════════
function addDonor(body) {
  if(!body.name)  throw new Error('Full name is required');
  if(!body.email) throw new Error('Email is required for notifications');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=getOrCreateSheet(ss,SHEETS.DONORS,['Donor ID','Donor Name','Email','Phone','Address','Registered On']);
  const id=generateId('DON',ss,SHEETS.DONORS,'Donor ID');
  ws.appendRow([id,String(body.name).trim(),String(body.email).trim().toLowerCase(),body.phone||'',body.address||'',fmt(new Date(),'yyyy-MM-dd')]);

  const email=String(body.email).trim().toLowerCase();
  sendEmail(email,
    '🎉 Welcome to Grocery Donation Drive — Donor ID: '+id,
    'Hi '+body.name+',\n\nThank you for registering as a donor!\n\nYour Donor ID: '+id+'\n\nPlease save this ID — you will need it each time you make a donation.\n\nYou will receive email notifications when your donations are collected, and when the items you donated are delivered to recipients. We will also keep you updated on inventory needs and community impact.\n\nThank you for your generosity!\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL);

  sendEmail(ADMIN_EMAIL,
    '[New Donor] '+body.name+' registered — '+id,
    'A new donor has registered.\n\nDonor ID: '+id+'\nName: '+body.name+'\nEmail: '+email+'\nPhone: '+(body.phone||'—')+'\nAddress: '+(body.address||'—')+'\nDate: '+new Date().toString());

  return {success:true,donorId:id,message:'Donor registered: '+id};
}

function addRecipient(body) {
  if(!body.name)  throw new Error('Full name is required');
  if(!body.email) throw new Error('Email is required for notifications');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=getOrCreateSheet(ss,SHEETS.RECIPIENTS,['Recipient ID','Recipient Name','Email','Phone','Address','Registered On']);
  const id=generateId('RCP',ss,SHEETS.RECIPIENTS,'Recipient ID');
  ws.appendRow([id,String(body.name).trim(),String(body.email).trim().toLowerCase(),body.phone||'',body.address||'',fmt(new Date(),'yyyy-MM-dd')]);

  const email=String(body.email).trim().toLowerCase();
  sendEmail(email,
    '🎉 Welcome to Grocery Donation Drive — Recipient ID: '+id,
    'Hi '+body.name+',\n\nYou have been successfully registered.\n\nYour Recipient ID: '+id+'\n\nPlease save this ID — you will need it each time you submit a request.\n\nYou can request available items through the Recipient Portal. You will receive email notifications when your requests are approved and when items are delivered. You can also request items not currently in our list using the Custom Request option.\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL);

  sendEmail(ADMIN_EMAIL,
    '[New Recipient] '+body.name+' registered — '+id,
    'A new recipient has registered.\n\nRecipient ID: '+id+'\nName: '+body.name+'\nEmail: '+email+'\nPhone: '+(body.phone||'—')+'\nAddress: '+(body.address||'—')+'\nDate: '+new Date().toString());

  return {success:true,recipientId:id,message:'Recipient registered: '+id};
}

// ════════════════════════════════════════════════
// ITEM MASTER
// ════════════════════════════════════════════════
function addItem(body) {
  if(!body.itemName) throw new Error('Missing itemName');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=getOrCreateSheet(ss,SHEETS.ITEMS,['Item Name','Category','Added On','Active']);
  const vals=ws.getDataRange().getValues();
  const exists=vals.slice(1).some(r=>String(r[0]).trim().toLowerCase()===String(body.itemName).trim().toLowerCase()&&String(r[3]).toLowerCase()!=='no');
  if(exists)return{success:false,error:'Item already exists'};
  ws.appendRow([String(body.itemName).trim(),body.category||'General',fmt(new Date(),'yyyy-MM-dd'),'Yes']);
  return{success:true,message:'Item added: '+body.itemName,data:getAllData()};
}

function deleteItem(body) {
  if(!body.itemName) throw new Error('Missing itemName');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=ss.getSheetByName(SHEETS.ITEMS);
  if(!ws)return{success:false,error:'Item Master sheet not found'};
  const vals=ws.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][0]).trim().toLowerCase()===String(body.itemName).trim().toLowerCase()){
      ws.getRange(i+1,4).setValue('No');
      return{success:true,message:'Item removed: '+body.itemName,data:getAllData()};
    }
  }
  return{success:false,error:'Item not found'};
}

// ════════════════════════════════════════════════
// CUSTOM ITEM REQUESTS
// ════════════════════════════════════════════════
function addCustomRequest(body) {
  if(!body.recipientId) throw new Error('Recipient ID is required');
  if(!body.itemName)    throw new Error('Item name is required');
  if(!body.notes)       throw new Error('Please explain why you need this item');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  verifyId(ss,SHEETS.RECIPIENTS,'Recipient ID',body.recipientId);
  const ws=getOrCreateSheet(ss,SHEETS.CUSTOM_REQUESTS,['Custom ID','Date','Recipient ID','Item Requested','Notes','Status','Admin Note','Resolved On']);
  const customId=generateId('CRQ',ss,SHEETS.CUSTOM_REQUESTS,'Custom ID');
  const now=new Date();
  ws.appendRow([customId,fmt(now,'yyyy-MM-dd'),String(body.recipientId).trim(),String(body.itemName).trim(),body.notes||'','Pending','','']);

  const recName=getPersonName(ss,SHEETS.RECIPIENTS,'Recipient ID',body.recipientId);
  sendEmail(ADMIN_EMAIL,
    '⚡ Action needed — Custom item request: '+body.itemName,
    'A recipient has requested an item not currently in your list. This is a priority request.\n\nCustom Request ID: '+customId+'\nRecipient ID: '+body.recipientId+'\nRecipient Name: '+recName+'\nItem Requested: '+body.itemName+'\nReason: '+body.notes+'\nDate: '+fmt(now,'yyyy-MM-dd HH:mm')+'\n\nPlease review in the admin portal under Custom Requests tab. Note: The recipient has explicitly asked for this item.\n\n'+ADMIN_EMAIL);

  return{success:true,customId,message:'Custom request submitted. The admin will review and email you.'};
}

function updateCustomRequest(body) {
  if(!body.rowIndex)  throw new Error('Missing rowIndex');
  if(!body.adminNote) throw new Error('Admin note is required');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=ss.getSheetByName(SHEETS.CUSTOM_REQUESTS);
  if(!ws) throw new Error('Custom Item Requests sheet not found');
  const allData=ws.getDataRange().getValues();
  const headers=allData[0].map(h=>String(h).trim());
  const rowNum=Number(body.rowIndex);
  const status=String(body.status||'');
  const note=String(body.adminNote||'');
  const rowArr=allData[rowNum-1];
  const rowObj={};headers.forEach((h,i)=>rowObj[h]=rowArr[i]);

  const setCol=(col,val)=>{const ci=headers.indexOf(col);if(ci>-1)ws.getRange(rowNum,ci+1).setValue(val);};
  setCol('Status',status);setCol('Admin Note',note);setCol('Resolved On',fmt(new Date(),'yyyy-MM-dd HH:mm'));

  const recipientId=String(rowObj['Recipient ID']||'');
  const itemName=String(rowObj['Item Requested']||'');
  const customId=String(rowObj['Custom ID']||'');
  const email=getPersonEmail(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
  const recName=getPersonName(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
  const isOk=status==='Fulfilled';

  // AUTO-ADD to Item Master when fulfilled so it appears in donor/recipient dropdowns
  if(isOk && itemName){
    const itemWs=getOrCreateSheet(ss,SHEETS.ITEMS,['Item Name','Category','Added On','Active']);
    const existingItems=itemWs.getDataRange().getValues();
    const alreadyExists=existingItems.slice(1).some(r=>
      String(r[0]).trim().toLowerCase()===itemName.toLowerCase()&&String(r[3]).toLowerCase()!=='no');
    if(!alreadyExists){
      itemWs.appendRow([itemName,'General',fmt(new Date(),'yyyy-MM-dd'),'Yes']);
      Logger.log('Auto-added to Item Master: '+itemName);
    }
  }

  const subject=isOk?'✅ Custom item request fulfilled — '+customId:'❌ Custom item request update — '+customId;
  const msgBody=isOk
    ?'Hi '+recName+',\n\nGreat news! Your request for "'+itemName+'" (ID: '+customId+') has been reviewed and the item has been added to our list. You can now request it through the Recipient Portal.\n\nAdmin note: '+note+'\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL
    :'Hi '+recName+',\n\nYour request for "'+itemName+'" (ID: '+customId+') has been reviewed but we are unable to fulfil it at this time.\n\nReason: '+note+'\n\nPlease contact us if you have any questions.\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL;

  if(email) sendEmail(email,subject,msgBody);
  sendEmail(ADMIN_EMAIL,'[Admin CC] '+subject,msgBody+'\n\nRecipient ID: '+recipientId);
  return{success:true,message:'Custom request marked as '+status,data:getAllData()};
}


// ════════════════════════════════════════════════
// WAITLIST
// ════════════════════════════════════════════════
function addWaitlist(body) {
  if(!body.recipientId) throw new Error('Recipient ID required');
  if(!body.item)        throw new Error('Item required');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  verifyId(ss,SHEETS.RECIPIENTS,'Recipient ID',body.recipientId);
  const ws=getOrCreateSheet(ss,SHEETS.WAITLIST,['Waitlist ID','Date Added','Recipient ID','Item','Delivery Notes','Status','Notified On']);
  const wid=generateId('WL',ss,SHEETS.WAITLIST,'Waitlist ID');
  const now=new Date();
  ws.appendRow([wid,fmt(now,'yyyy-MM-dd'),body.recipientId,body.item,body.notes||'','Waiting','']);

  const recEmail=getPersonEmail(ss,SHEETS.RECIPIENTS,'Recipient ID',body.recipientId);
  const recName=getPersonName(ss,SHEETS.RECIPIENTS,'Recipient ID',body.recipientId);
  // No email notifications for waitlist — admin portal visibility only
  Logger.log('Waitlist entry created: '+wid+' for '+body.recipientId+' item: '+body.item);
  return {success:true,waitlistId:wid,message:'Added to waitlist for '+body.item};
}


function updateWaitlistStatus(body) {
  if(!body.rowIndex) throw new Error('Missing rowIndex');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=ss.getSheetByName(SHEETS.WAITLIST);
  if(!ws) throw new Error('Waitlist sheet not found');
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const statusCol=headers.indexOf('Status');
  const notifiedCol=headers.indexOf('Notified On');
  const rowNum=Number(body.rowIndex);
  if(statusCol>-1) ws.getRange(rowNum,statusCol+1).setValue(body.status||'Cancelled');
  if(notifiedCol>-1) ws.getRange(rowNum,notifiedCol+1).setValue(fmt(new Date(),'yyyy-MM-dd HH:mm'));
  return {success:true,message:'Waitlist entry updated',data:getAllData()};
}


// ── Find the specific donation that funded the most recent delivery of an item ──
// Logic: FIFO consumption — take collected donations in order of collection date,
// subtract delivered requests in order, find which donation covers the current delivery
function findFundingDonation(ss, itemName) {
  const donWs=ss.getSheetByName(SHEETS.DONATIONS);
  const reqWs=ss.getSheetByName(SHEETS.REQUESTS);
  if(!donWs||!reqWs)return'';

  const donData=donWs.getDataRange().getValues();
  const donH=donData[0].map(h=>String(h).trim());
  const di=donH.indexOf('Item'),ds=donH.indexOf('Status'),dd=donH.indexOf('Donation ID'),dc=donH.indexOf('Collected On'),dq=donH.indexOf('Quantity');

  // All collected donations for this item, sorted by collection date ascending (FIFO)
  const collectedDons=donData.slice(1)
    .filter(r=>String(r[di]||'').trim()===itemName&&String(r[ds]||'').toLowerCase()==='collected')
    .map(r=>({id:String(r[dd]||''),date:String(r[dc]||''),qty:Number(r[dq])||0}))
    .sort((a,b)=>a.date.localeCompare(b.date));
  if(!collectedDons.length)return'';

  const reqData=reqWs.getDataRange().getValues();
  const reqH=reqData[0].map(h=>String(h).trim());
  const ri=reqH.indexOf('Item Requested'),rs=reqH.indexOf('Status'),rq=reqH.indexOf('Qty Requested'),rdo=reqH.indexOf('Delivered On');

  // All delivered requests for this item, sorted by delivery date ascending (FIFO consumption)
  const deliveredReqs=reqData.slice(1)
    .filter(r=>String(r[ri]||'').trim()===itemName&&String(r[rs]||'').toLowerCase()==='delivered')
    .map(r=>({qty:Number(r[rq])||0,deliveredOn:String(r[rdo]||'')}))
    .sort((a,b)=>a.deliveredOn.localeCompare(b.deliveredOn));

  // Walk through donations FIFO, consume against deliveries, find which donation
  // the LAST delivered request consumed from
  let donIdx=0,remaining=collectedDons[0]?collectedDons[0].qty:0;
  let fundingDonId=collectedDons[0]?collectedDons[0].id:'';

  for(let i=0;i<deliveredReqs.length;i++){
    let need=deliveredReqs[i].qty;
    while(need>0&&donIdx<collectedDons.length){
      const take=Math.min(need,remaining);
      need-=take;remaining-=take;
      fundingDonId=collectedDons[donIdx].id; // track which donation is being consumed
      if(remaining<=0){
        donIdx++;
        if(donIdx<collectedDons.length) remaining=collectedDons[donIdx].qty;
      }
    }
  }
  return fundingDonId;
}

function notifyWaitlist(ss, itemName, fulfilledByDonationId) {
  // Called when a request is Delivered or Rejected — marks next FIFO waitlist entry
  // No emails — admin portal visibility only
  const ws=ss.getSheetByName(SHEETS.WAITLIST);
  if(!ws)return;
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const itemCol=headers.indexOf('Item');
  const statusCol=headers.indexOf('Status');
  const ridCol=headers.indexOf('Recipient ID');
  const notifiedCol=headers.indexOf('Notified On');
  const dateCol=headers.indexOf('Date Added');
  // Add Donation Txn ID column if missing
  let donTxnCol=headers.indexOf('Fulfilled By Donation');
  if(donTxnCol<0){
    const lastCol=ws.getLastColumn();
    ws.getRange(1,lastCol+1).setValue('Fulfilled By Donation');
    // Re-read headers so index is accurate (1-indexed correction)
    const freshHeaders=ws.getDataRange().getValues()[0].map(h=>String(h).trim());
    donTxnCol=freshHeaders.indexOf('Fulfilled By Donation');
  }

  // Find oldest (FIFO) Waiting entry for this item — sort by Date Added ascending
  let candidates=[];
  for(let i=1;i<data.length;i++){
    if(String(data[i][itemCol]||'').trim()===itemName && String(data[i][statusCol]||'').toLowerCase()==='waiting'){
      candidates.push({row:i+1,date:data[i][dateCol]||'9999'});
    }
  }
  if(!candidates.length)return;
  // Sort oldest first (FIFO)
  candidates.sort((a,b)=>String(a.date).localeCompare(String(b.date)));
  const targetRow=candidates[0].row;

  // Mark as Completed (not just Notified — this is the final state)
  ws.getRange(targetRow,statusCol+1).setValue('Completed');
  ws.getRange(targetRow,notifiedCol+1).setValue(fmt(new Date(),'yyyy-MM-dd HH:mm'));
  if(fulfilledByDonationId){
    ws.getRange(targetRow,donTxnCol+1).setValue(fulfilledByDonationId);
  }
  Logger.log('Waitlist completed: row '+targetRow+' for '+itemName+(fulfilledByDonationId?' by '+fulfilledByDonationId:''));
}

// ════════════════════════════════════════════════
// NOTICE BOARD
// ════════════════════════════════════════════════
function addNotice(body) {
  if(!body.notice) throw new Error('Notice text is required');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=getOrCreateSheet(ss,SHEETS.NOTICES,['Date','Notice','Active','Added By']);
  ws.appendRow([fmt(new Date(),'yyyy-MM-dd'),String(body.notice).trim(),'Yes','Admin']);
  return{success:true,message:'Notice posted',data:getAllData()};
}

function deleteNotice(body) {
  if(!body.rowIndex) throw new Error('Missing rowIndex');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ws=ss.getSheetByName(SHEETS.NOTICES);
  if(!ws)return{success:false,error:'Notice Board sheet not found'};
  const headers=ws.getDataRange().getValues()[0].map(h=>String(h).trim());
  const activeCol=headers.indexOf('Active');
  if(activeCol>-1)ws.getRange(Number(body.rowIndex),activeCol+1).setValue('No');
  return{success:true,message:'Notice removed',data:getAllData()};
}

// ════════════════════════════════════════════════
// EMAIL FUNCTIONS
// ════════════════════════════════════════════════
// ── EMAIL PAUSED — set EMAIL_PAUSED=false to re-enable ──
const EMAIL_PAUSED = true;
function sendEmail(to,subject,body){
  if(!to)return;
  if(EMAIL_PAUSED){
    Logger.log('[EMAIL PAUSED] To: '+to+' | Subject: '+subject);
    return;
  }
  MailApp.sendEmail({to:String(to),subject:String(subject),body:String(body)});
}

function emailDonationStatus(ss,rowObj,status,comment){
  const donorId=String(rowObj['Donor ID']||'');
  const donationId=String(rowObj['Donation ID']||donorId);
  const item=String(rowObj['Item']||'—');
  const qty=rowObj['Quantity']||0;
  const date=String(rowObj['Date']||'—');
  const email=getPersonEmail(ss,SHEETS.DONORS,'Donor ID',donorId);
  const name=getPersonName(ss,SHEETS.DONORS,'Donor ID',donorId);
  const commentLine=comment?'\n\nAdmin comment: '+comment:'';

  let subject,body;
  if(status==='Collected'){
    subject='✅ Donation collected — '+donationId;
    body='Hi '+name+',\n\nYour donation of '+qty+' × '+item+' (ID: '+donationId+') submitted on '+date+' has been collected. Thank you for your generosity!'+commentLine+'\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL;
  }else if(status==='Rejected'){
    subject='❌ Donation could not be collected — '+donationId;
    body='Hi '+name+',\n\nUnfortunately your donation of '+qty+' × '+item+' (ID: '+donationId+') could not be collected at this time.'+commentLine+'\n\nPlease contact us for more information.\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL;
  }else{
    subject='Donation update: '+status+' — '+donationId;
    body='Hi '+name+',\n\nYour donation '+donationId+' ('+qty+' × '+item+') status: '+status+'.'+commentLine+'\n\nGrocery Donation Tracker';
  }
  if(email)sendEmail(email,subject,body);
  sendEmail(ADMIN_EMAIL,'[Admin CC] '+subject,body+'\n\n---\nDonor ID: '+donorId+'\nEmail: '+(email||'Not registered'));
}

function emailRequestStatus(ss,rowObj,status,comment,eta){
  const recipientId=String(rowObj['Recipient ID']||'');
  const requestId=String(rowObj['Request ID']||recipientId);
  const item=String(rowObj['Item Requested']||'—');
  const qty=rowObj['Qty Requested']||1;
  const date=String(rowObj['Request Date']||'—');
  const email=getPersonEmail(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
  const name=getPersonName(ss,SHEETS.RECIPIENTS,'Recipient ID',recipientId);
  const commentLine=comment?'\n\nAdmin comment: '+comment:'';
  const etaLine=eta?'\n\nEstimated delivery: '+eta:'';

  const msgMap={
    'Approved'        :['✅ Request approved — '+requestId,'Hi '+name+',\n\nYour request for '+qty+' × '+item+' (ID: '+requestId+') has been approved.'+etaLine+'\n\nWe will arrange delivery to your address. Please ensure someone is available to receive it.'+commentLine+'\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL],
    // Delivery Pending status removed — workflow is now Pending → Approved → Delivered
    'Delivered'       :['📦 Delivered — '+requestId,'Hi '+name+',\n\nYour request for '+qty+' × '+item+' (ID: '+requestId+') has been delivered.\n\nPlease take time to write a feedback in the Recipient Portal — it only takes a minute and means a lot to our donors!'+commentLine+'\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL],
    'Rejected'        :['❌ Request rejected — '+requestId,'Hi '+name+',\n\nYour request for '+qty+' × '+item+' (ID: '+requestId+') could not be fulfilled at this time.'+commentLine+'\n\nPlease contact us for more information.\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL],
  };

  const[subject,body]=msgMap[status]||['Request update: '+status+' — '+requestId,'Your request '+requestId+' status: '+status+'.'+commentLine+'\n\nGrocery Donation Tracker'];
  if(email)sendEmail(email,subject,body);
  sendEmail(ADMIN_EMAIL,'[Admin CC] '+subject,body+'\n\n---\nRecipient ID: '+recipientId+'\nEmail: '+(email||'Not registered'));
}

function emailAllDonors_NewRequest(ss,itemName,recipientName,requestId){
  if(EMAIL_PAUSED){Logger.log('[EMAIL PAUSED] emailAllDonors_NewRequest: '+itemName);return;}
  const emails=getAllDonorEmails(ss);
  if(!emails.length)return;
  MailApp.sendEmail({to:ADMIN_EMAIL,bcc:emails.join(','),
    subject:'🛒 Community needs '+itemName+' — consider donating',
    body:'Hi,\n\nA student in our community has just submitted a request for:\n\n  Item: '+itemName+'\n  Request ID: '+requestId+'\n\nIf you have this item available, please consider donating through the Donor Portal. Every contribution directly helps a student in need.\n\nThank you for your generosity!\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL});
}

function emailAllDonors_LowStock(ss,lowItems,customRequests){
  if(EMAIL_PAUSED){Logger.log('[EMAIL PAUSED] emailAllDonors_LowStock');return;}
  const emails=getAllDonorEmails(ss);
  if(!emails.length)return;

  // Get ALL items from Item Master (not just those in inventory)
  // Items never donated yet count as 0 balance = also low
  const itemWs=ss.getSheetByName(SHEETS.ITEMS);
  let allItems=[];
  if(itemWs){
    const data=itemWs.getDataRange().getValues();
    const headers=data[0].map(h=>String(h).trim());
    const nameCol=headers.indexOf('Item Name');
    const activeCol=headers.indexOf('Active');
    allItems=data.slice(1)
      .filter(r=>String(r[activeCol]||'').toLowerCase()!=='no')
      .map(r=>String(r[nameCol]||'').trim())
      .filter(Boolean);
  }

  // Build balance map from inventory
  const balMap={};
  lowItems.forEach(i=>balMap[i.item]=i.balance);

  // Find items either in lowItems OR never donated (balance effectively 0)
  const needDonation=allItems.filter(name=>{
    const bal=balMap[name]??null;
    return bal===null||(bal>=0&&bal<LOW_STOCK_THRESHOLD);
  });

  if(!needDonation.length&&!lowItems.length)return;

  const pending=(customRequests||[]).filter(r=>(r['Status']||'').toLowerCase()==='pending');
  const itemList=needDonation.map(name=>{
    const bal=balMap[name]??0;
    return'  • '+name+(bal===0?': none available':': only '+bal+' left');
  }).join('\n')||'  (see custom requests below)';

  let customSection='';
  if(pending.length>0){
    customSection='\n\nPRIORITY — Recipients have specifically requested these items:\n'
      +pending.map(r=>'  • '+r['Item Requested']+' — "'+r['Notes']+'"').join('\n')
      +'\n\nThese recipients have explicitly asked for these items.';
  }

  MailApp.sendEmail({to:ADMIN_EMAIL,bcc:emails.join(','),
    subject:'⚠️ Items running low — your donation is appreciated',
    body:'Hi,\n\nSome items in our grocery drive need replenishment:\n\n'+itemList
      +customSection
      +'\n\nIf you can donate any of these items, please visit the Donor Portal.'
      +'\n\nThank you!\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL});
}

// ════════════════════════════════════════════════
// DAILY ADMIN REMINDER (10am PST)
// ════════════════════════════════════════════════
function dailyAdminReminderTrigger(){
  if(EMAIL_PAUSED){Logger.log('[EMAIL PAUSED] dailyAdminReminderTrigger');return;}
  const data=getAllData();
  const don=data.donations||[],req=data.requests||[];
  const pendingColl=don.filter(r=>String(r['Status']||'submitted').toLowerCase()==='submitted');
  const pendingReq=req.filter(r=>String(r['Status']||'').toLowerCase()==='pending');
  const dpReq=req.filter(r=>String(r['Status']||'').toLowerCase()==='delivery pending');
  const pendingCustom=(data.customRequests||[]).filter(r=>String(r['Status']||'').toLowerCase()==='pending');

  if(pendingColl.length===0&&pendingReq.length===0&&dpReq.length===0&&pendingCustom.length===0)return;

  let body='Good morning! Here is your daily action summary for the Grocery Donation Tracker.\n\n';
  if(pendingColl.length>0){
    body+='📦 DONATIONS TO COLLECT ('+pendingColl.length+'):\n';
    body+=pendingColl.map(r=>'  • '+r['Donation ID']+' | '+r['Item']+' ×'+r['Quantity']+' | Submitted: '+String(r['Submitted On']||r['Date']).split(' ')[0]).join('\n')+'\n\n';
  }
  if(pendingReq.length>0){
    body+='⏳ REQUESTS TO APPROVE ('+pendingReq.length+'):\n';
    body+=pendingReq.map(r=>'  • '+r['Request ID']+' | '+r['Item Requested']+' | Pending since: '+String(r['Pending On']||r['Request Date']).split(' ')[0]).join('\n')+'\n\n';
  }
  if(dpReq.length>0){
    body+='🚚 OUT FOR DELIVERY ('+dpReq.length+') — confirm delivered:\n';
    body+=dpReq.map(r=>'  • '+r['Request ID']+' | '+r['Item Requested']+' | ETA: '+(r['ETA']||'not set')).join('\n')+'\n\n';
  }
  if(pendingCustom.length>0){
    body+='🔍 CUSTOM REQUESTS PENDING ('+pendingCustom.length+'):\n';
    body+=pendingCustom.map(r=>'  • '+r['Custom ID']+' | '+r['Item Requested']+' | "'+r['Notes']+'"').join('\n')+'\n\n';
  }
  body+='Please log in to the admin portal to take action.\n\nGrocery Donation Tracker\n'+ADMIN_EMAIL;
  sendEmail(ADMIN_EMAIL,'⚡ Daily action summary — Grocery Donation Tracker',body);
}

// ════════════════════════════════════════════════
// WEEKLY LOW STOCK ALERT (Monday 10am PST)
// ════════════════════════════════════════════════
function weeklyLowStockTrigger(){
  if(EMAIL_PAUSED){Logger.log('[EMAIL PAUSED] weeklyLowStockTrigger');return;}
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const data=getAllData();
  const lowItems=data.inventory.filter(i=>i.balance>=0&&i.balance<LOW_STOCK_THRESHOLD);
  // Always run — includes items never donated (balance = 0)
  emailAllDonors_LowStock(ss,lowItems,data.customRequests||[]);
}

// ════════════════════════════════════════════════
// WEEKLY ADMIN REPORT (Monday)
// ════════════════════════════════════════════════
function weeklyReportTrigger(){
  if(EMAIL_PAUSED){Logger.log('[EMAIL PAUSED] weeklyReportTrigger');return;}
  const data=getAllData();
  const now=new Date();
  const don=data.donations||[],req=data.requests||[];
  const donors=data.donors||[],rec=data.recipients||[];
  const inv=data.inventory||[];
  const collected=don.filter(r=>String(r['Status']||'').toLowerCase()==='collected').length;
  const pendColl=don.filter(r=>String(r['Status']||'submitted').toLowerCase()==='submitted').length;
  const delivered=req.filter(r=>String(r['Status']||'').toLowerCase()==='delivered').length;
  const pendReq=req.filter(r=>String(r['Status']||'').toLowerCase()==='pending').length;
  const dp=req.filter(r=>String(r['Status']||'').toLowerCase()==='delivery pending').length;
  const lowItems=inv.filter(i=>i.balance<LOW_STOCK_THRESHOLD);
  const withFb=req.filter(r=>r['Feedback']);

  const invLines=inv.map(i=>'  • '+i.item+': '+i.donated+' donated, '+i.distributed+' distributed, '+i.balance+' available').join('\n')||'  No inventory data.';
  const lowLines=lowItems.length?lowItems.map(i=>'  ⚠ '+i.item+': '+i.balance+' remaining').join('\n'):'  All items well stocked.';
  const fbLines=withFb.slice(-3).map(r=>'  "'+r['Feedback']+'"').join('\n\n')||'  None this week.';

  const subject='[Weekly Admin Report] Grocery Donation Tracker — '+fmt(now,'dd MMM yyyy');
  const body=[
    'GROCERY DONATION TRACKER — WEEKLY ADMIN SUMMARY',
    'Week of: '+fmt(now,'dd MMM yyyy')+'  |  Generated: '+fmt(now,'yyyy-MM-dd HH:mm'),
    '─'.repeat(52),'',
    'OVERVIEW',
    '  Donors: '+donors.length+'  |  Recipients: '+rec.length,
    '  Donations: '+don.length+' ('+collected+' collected, '+pendColl+' awaiting collection)',
    '  Requests: '+req.length+' ('+delivered+' delivered, '+pendReq+' pending, '+dp+' out for delivery)',
    '','LIVE INVENTORY',invLines,'','LOW STOCK (< '+LOW_STOCK_THRESHOLD+')',lowLines,
    '','RECENT FEEDBACK FROM RECIPIENTS',fbLines,
    '','─'.repeat(52),'Grocery Donation Tracker · '+ADMIN_EMAIL,
  ].join('\n');
  sendEmail(ADMIN_EMAIL,subject,body);
}

// ════════════════════════════════════════════════
// WEEKLY DONOR REPORT (Friday)
// ════════════════════════════════════════════════
function weeklyDonorReportTrigger(){
  if(EMAIL_PAUSED){Logger.log('[EMAIL PAUSED] weeklyDonorReportTrigger');return;}
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const data=getAllData();
  const now=new Date();
  const don=data.donations||[],req=data.requests||[];
  const inv=data.inventory||[];
  const cr=(data.customRequests||[]).filter(r=>String(r['Status']||'').toLowerCase()==='pending');
  const collected=don.filter(r=>String(r['Status']||'').toLowerCase()==='collected');
  const delivered=req.filter(r=>String(r['Status']||'').toLowerCase()==='delivered');
  const withFb=req.filter(r=>r['Feedback']);
  const invLines=inv.length?inv.map(i=>'  • '+i.item+': '+i.balance+' available').join('\n'):'  No items in inventory yet.';
  const crSection=cr.length?'\n\nSPECIAL REQUESTS FROM RECIPIENTS\nThese items have been explicitly requested by community members:\n'+cr.map(r=>'  • '+r['Item Requested']+' (Reason: '+r['Notes']+')').join('\n')+'\n\nIf you can donate any of these, please visit the Donor Portal.':'';
  const fbSection=withFb.length?'\n\nWHAT YOUR DONATIONS HAVE MEANT:\n'+withFb.slice(-3).map(r=>'"'+r['Feedback']+'"').join('\n\n'):'';

  const emails=getAllDonorEmails(ss);
  if(!emails.length)return;

  const subject='💚 Your impact this week — Grocery Donation Drive Update';
  const body='Dear Donor,\n\nThank you for being part of our Grocery Donation Drive. Here is what your generosity has helped achieve this week:\n\n'
    +'📦 Total units donated (collected): '+collected.reduce((s,r)=>s+(parseInt(r['Quantity'])||0),0)+'\n'
    +'📋 Total requests made: '+req.length+'\n'
    +'✅ Items successfully delivered: '+delivered.length+'\n\n'
    +'CURRENT INVENTORY\n'+invLines
    +crSection+fbSection
    +'\n\nIf you have items to donate, we always appreciate your continued support. Visit the Donor Portal at any time to make a new donation.\n\nWith gratitude,\nGrocery Donation Tracker\n'+ADMIN_EMAIL;

  MailApp.sendEmail({to:ADMIN_EMAIL,bcc:emails.join(','),subject,body});
}

// ════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════
function getAllDonorEmails(ss){
  const ws=ss.getSheetByName(SHEETS.DONORS);if(!ws)return[];
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const emCol=headers.indexOf('Email');if(emCol<0)return[];
  return data.slice(1).map(r=>String(r[emCol]||'').trim().toLowerCase()).filter(e=>e&&e.includes('@'));
}

function getPersonEmail(ss,sheetName,idField,idValue){
  const ws=ss.getSheetByName(sheetName);if(!ws)return null;
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const idCol=headers.indexOf(idField),emCol=headers.indexOf('Email');
  if(idCol<0||emCol<0)return null;
  for(let i=1;i<data.length;i++){
    if(String(data[i][idCol]).trim()===String(idValue).trim())return String(data[i][emCol]||'').trim().toLowerCase()||null;
  }
  return null;
}

function getPersonName(ss,sheetName,idField,idValue){
  const ws=ss.getSheetByName(sheetName);if(!ws)return idValue;
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const idCol=headers.indexOf(idField);
  const nmF=sheetName===SHEETS.DONORS?'Donor Name':'Recipient Name';
  const nmCol=headers.indexOf(nmF);
  if(idCol<0||nmCol<0)return idValue;
  for(let i=1;i<data.length;i++){
    if(String(data[i][idCol]).trim()===String(idValue).trim())return String(data[i][nmCol]||'').trim()||idValue;
  }
  return idValue;
}

function verifyId(ss,sheetName,idField,idValue){
  const ws=ss.getSheetByName(sheetName);if(!ws)return;
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const idCol=headers.indexOf(idField);
  const found=data.slice(1).some(r=>String(r[idCol]).trim()===String(idValue).trim());
  if(!found)throw new Error(idField.replace(' ID','')+' ID "'+idValue+'" not found. Please register first.');
}

function logHistory(ss,type,personId,item,qty,ts,status,txnId){
  const ws=getOrCreateSheet(ss,SHEETS.HISTORY,['Timestamp','Type','Transaction ID','Person ID','Item','Quantity','Status']);
  ws.appendRow([fmt(ts,'yyyy-MM-dd HH:mm:ss'),type,txnId||'',personId,item,Number(qty),status||'']);
}

function getOrCreateSheet(ss,name,headers){
  let ws=ss.getSheetByName(name);
  if(!ws){ws=ss.insertSheet(name);ws.appendRow(headers);}
  return ws;
}

function generateId(prefix,ss,sheetName,idCol){
  const ws=ss.getSheetByName(sheetName);if(!ws||ws.getLastRow()<2)return prefix+'001';
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const ci=headers.indexOf(idCol);
  const nums=data.slice(1).map(r=>String(r[ci]||'')).filter(id=>id.startsWith(prefix)).map(id=>parseInt(id.replace(prefix,''))||0);
  return prefix+String(nums.length?Math.max(...nums)+1:1).padStart(3,'0');
}

function generateTxnId(personId,suffix,ss,sheetName,txnCol){
  const ws=ss.getSheetByName(sheetName);
  const prefix=personId+'_'+suffix;
  if(!ws||ws.getLastRow()<2)return prefix+'001';
  const data=ws.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const ci=headers.indexOf(txnCol);
  const nums=data.slice(1).map(r=>String(r[ci]||'')).filter(id=>id.startsWith(prefix)).map(id=>parseInt(id.split('_'+suffix).pop())||0);
  return prefix+String(nums.length?Math.max(...nums)+1:1).padStart(3,'0');
}

function fmt(date,pattern){return Utilities.formatDate(date,Session.getScriptTimeZone(),pattern);}
