const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  INV: 'Inventory',
  LOGS: 'Logs',
  USERS: 'Users',
  REQ: 'Requests',
  DISCREP: 'Discrepancies',
  DROPS: 'Dropdowns',
  PO_DB: 'PO Database',
  PO_ASSIGN: 'PO Assignments',
  NOTIF: 'Notifications'
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.scriptUrl = ScriptApp.getService().getUrl();
  template.logoUrl = getLogoSafe(); 
  
  return template.evaluate()
    .setTitle('Inventory System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData(userProfile) {
  try {
    if (!SS) throw new Error("Could not access spreadsheet.");
    const iSheet = SS.getSheetByName(SHEETS.INV);
    const lSheet = SS.getSheetByName(SHEETS.LOGS);
    const uSheet = SS.getSheetByName(SHEETS.USERS);
    const rSheet = SS.getSheetByName(SHEETS.REQ);
    const dSheet = SS.getSheetByName(SHEETS.DROPS) || SS.insertSheet(SHEETS.DROPS);
    const poSheet = SS.getSheetByName(SHEETS.PO_DB);
    SpreadsheetApp.flush();
    
    let inventory = [];
    if (iSheet && iSheet.getLastRow() > 0) {
      const invData = iSheet.getDataRange().getValues();
      const headers = invData[0];
      
      for (let i = 1; i < invData.length; i++) {
        if (!invData[i][0] || invData[i][0].toString().trim() === "") continue;
        let item = {
          code: invData[i][0].toString(),
          name: invData[i][1].toString(),
          uom: invData[i][2]?.toString() || '',
          balances: {}
        };
        for (let j = 3; j < headers.length; j++) {
           if (headers[j] && headers[j].toString().trim() !== "") {
              item.balances[headers[j]] = Number(invData[i][j]) || 0;
           }
        }
        inventory.push(item);
      }
    }

    let dropdowns = { clients: [], locations: [], siteIds: [], sites: [], siteMap: {}, siteDetails: {}, pos: [], pendingPOs: [] };
    if (dSheet && dSheet.getLastRow() > 1) {
      const dData = dSheet.getDataRange().getValues();
      for (let i = 1; i < dData.length; i++) {
         if (dData[i][0]) dropdowns.clients.push(dData[i][0].toString());
         if (dData[i][1]) dropdowns.locations.push(dData[i][1].toString());
         if (dData[i][2]) dropdowns.siteIds.push(dData[i][2].toString());
         if (dData[i][3]) {
            let sName = dData[i][3].toString();
            let client = dData[i][0] ? dData[i][0].toString() : '';
            let sId = dData[i][2] ? dData[i][2].toString() : '';
            let mappedLoc = dData[i][4] ? dData[i][4].toString() : (dData[i][1] ? dData[i][1].toString() : '');
            let mappedWbs = dData[i][5] ? dData[i][5].toString() : ''; 

            dropdowns.sites.push(sName);
            if (mappedLoc) dropdowns.siteMap[sName] = mappedLoc;
            
            dropdowns.siteDetails[sName] = { client: client, siteId: sId, location: mappedLoc, wbs: mappedWbs };
         }
      }
    }

    if (poSheet && poSheet.getLastRow() > 1) {
       const poData = poSheet.getDataRange().getValues();
       for (let i = 1; i < poData.length; i++) {
           let po = poData[i][0] ? poData[i][0].toString().trim() : '';
           if (po && !dropdowns.pos.includes(po)) dropdowns.pos.push(po);
       }
    }

    const paSheet = SS.getSheetByName(SHEETS.PO_ASSIGN);
    let pendingPOMap = {};
    if (paSheet && paSheet.getLastRow() > 1) {
      const paData = paSheet.getRange(2, 1, paSheet.getLastRow() - 1, 11).getValues();
      paData.forEach(r => {
        let status = r[10] ? r[10].toString().trim() : '';
        if (status !== 'Pending') return;
        let doc = r[1] ? r[1].toString().trim() : '';
        if (!doc) return;
        if (!pendingPOMap[doc]) {
          pendingPOMap[doc] = {
            docNumber: doc,
            timestamp: r[0] && r[0].toISOString ? r[0].toISOString() : (r[0] || ''),
            user: r[2] ? r[2].toString() : '',
            location: r[3] ? r[3].toString() : '',
            siteName: r[4] ? r[4].toString() : '',
            items: []
          };
        }
        pendingPOMap[doc].items.push({
          code: r[5] ? r[5].toString() : '',
          name: r[6] ? r[6].toString() : '',
          uom: r[7] ? r[7].toString() : '',
          qty: Number(r[8]) || 0
        });
      });
    }
    dropdowns.pendingPOs = Object.values(pendingPOMap);

    let pending = [];
    let pendingDRs = [];
    if (rSheet && rSheet.getLastRow() > 1 && userProfile) {
      const reqData = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 13).getValues();
      
      let locList = userProfile.locAccess ? userProfile.locAccess.toString().split(',').map(s => s.trim()) : [];
      let siteList = userProfile.siteAccess ? userProfile.siteAccess.toString().split(',').map(s => s.trim()) : [];

      reqData.forEach(r => {
        let status = r[11];
        let reqAction = r[4];
        let reqLoc = r[5] ? r[5].toString().trim() : '';
        let reqSite = r[6] ? r[6].toString().trim() : '';
        let isMyTurn = false;
        
        if (userProfile.role === 'warehouseman') {
           let hasLocAccess = locList.length === 0 || locList.includes(reqLoc);
           if (status === 'Pending Return' && reqAction === 'RETURN_WH' && hasLocAccess) isMyTurn = true;
           if (status === 'Pending Receipt' && reqAction === 'TRANSFER_WH' && hasLocAccess) isMyTurn = true;
           
           if (status === 'Pending DR' && reqAction === 'DR_CREATE') {
               if (hasLocAccess) {
                  // Converted drId to string to prevent numerical bugs
                  pendingDRs.push({ drId: r[0].toString(), itemCode: r[7], itemName: r[8], unit: r[9], qty: r[10], location: reqLoc, sitename: reqSite, poNumber: "" });
               }
           }
        }
        else if (userProfile.role === 'team leader') {
           if (status === 'In Transit' && reqAction === 'ISSUE') {
              let hasSiteAccess = siteList.length === 0 || siteList.includes(reqSite);
              if (hasSiteAccess) isMyTurn = true;
           }
        }
        else if (userProfile.role === 'admin') {
           isMyTurn = true;
        }

        if (isMyTurn) {
          pending.push({
            id: r[0], date: Utilities.formatDate(r[1], "GMT+8", "yyyy-MM-dd HH:mm"), 
            requestor: r[2], role: r[3], action: reqAction, location: reqLoc, sitename: reqSite, 
            itemCode: r[7], itemName: r[8], unit: r[9], qty: r[10], status: r[11], remarks: r[12]
          });
        }
      });
    }

    let logs = [];
    let receiptDocs = {};
    if (lSheet && lSheet.getLastRow() > 1) {
      const logCols = Math.max(17, lSheet.getLastColumn());
      const logValues = lSheet.getRange(2, 1, lSheet.getLastRow() - 1, logCols).getValues();

      // Attach the PO Number captured at DR encode time to each pending DR (read-only display on receive)
      pendingDRs.forEach(pdr => {
         let matchedLog = logValues.find(l => l[1].toString() === pdr.drId && l[2] === 'DR_CREATE');
         if (matchedLog && matchedLog[17]) pdr.poNumber = matchedLog[17].toString();
      });

      logValues.forEach(l => {
          let docNum = l[1]?.toString();
          let action = l[2];
          let site = l[6];
          let loc = l[7];
          let code = l[8];
          let name = l[9];
          let uom = l[10];
          let qty = Number(l[12]) || 0;

          if (action === 'RECEIVE_DR') {
              // Captured site property so Return to Client can auto-populate it properly
              if (!receiptDocs[docNum]) receiptDocs[docNum] = { location: loc, site: site, items: {} };
              if (!receiptDocs[docNum].items[code]) receiptDocs[docNum].items[code] = { name: name, uom: uom, remaining: 0 };
              receiptDocs[docNum].items[code].remaining += qty;
          } else if (action === 'RETURN_CLIENT' || action === 'ISSUE' || action === 'TRANSFER_WH') {
              if (action === 'RETURN_CLIENT') {
                let suffix = docNum.substring(docNum.indexOf('-'));
                let possibleOrig = Object.keys(receiptDocs).find(k => k.endsWith(suffix));
                if (possibleOrig && receiptDocs[possibleOrig].items[code]) {
                    receiptDocs[possibleOrig].items[code].remaining -= qty;
                }
              }
          }
      });

      for (let doc in receiptDocs) {
          let hasItems = false;
          for (let code in receiptDocs[doc].items) {
              if (receiptDocs[doc].items[code].remaining > 0) hasItems = true;
          }
          if (!hasItems) delete receiptDocs[doc];
      }

      logs = logValues.reverse().map(l => {
        return [
          l[0].toISOString ? l[0].toISOString() : l[0],
          l[1]||'', l[2]||'', l[3]||'', l[4]||'', l[5]||'', l[6]||'', l[7]||'', l[8]||'', l[9]||'', l[10]||'',
          l[11]||'', l[12]||0, l[13]||0, l[14]||0, l[15]||'Completed', l[16]||'', l[17]||''
        ];
      });
    }
    
    return { inventory, logs, dropdowns, pending, pendingDRs, receiptDocs };
  } catch (e) {
    console.error("Error in getAppData: " + e.toString());
    throw e; 
  }
}

function processBulkTransaction(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (!SS) throw new Error("Could not access spreadsheet.");

    // PO is captured when a Delivery Receipt is encoded / purchase is logged. Either pick one or mark "PO to follow".
    if ((payload.action === 'DR_CREATE' || payload.action === 'PURCHASE_LOG') && !payload.poToFollow && !payload.poNumber) {
      throw new Error("PO Number is required (or mark 'PO to follow').");
    }

    const iSheet = SS.getSheetByName(SHEETS.INV);
    const lSheet = SS.getSheetByName(SHEETS.LOGS);
    const rSheet = SS.getSheetByName(SHEETS.REQ);
    
    let dSheet = SS.getSheetByName(SHEETS.DISCREP);
    if (!dSheet) {
      dSheet = SS.insertSheet(SHEETS.DISCREP);
      dSheet.appendRow(['Timestamp', 'Doc ID', 'User', 'Warehouse Location', 'Site Name', 'Item Code', 'Item Description', 'System Qty', 'Actual Returned Qty']);
      dSheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#0f172a").setFontColor("white");
      dSheet.setFrozenRows(1);
    }
    
    SpreadsheetApp.flush();
    const invRange = iSheet.getDataRange();
    let invData = invRange.getValues();
    
    if (invData.length === 0) invData = [['Item Code', 'Item Description', 'UOM']];
    let headers = invData[0];
    
    let locColIdx = -1; let siteColIdx = -1;
    if (payload.location && payload.location !== "-") {
      locColIdx = headers.indexOf(payload.location);
      if (locColIdx === -1) { headers.push(payload.location); locColIdx = headers.length - 1; }
    }
    if (payload.siteName && payload.siteName !== "-") {
      siteColIdx = headers.indexOf(payload.siteName);
      if (siteColIdx === -1) { headers.push(payload.siteName); siteColIdx = headers.length - 1; }
    }
    
    const d = new Date();
    let finalDocId = payload.refDoc ? payload.refDoc : "DOC-" + d.getFullYear() + ("0"+(d.getMonth()+1)).slice(-2) + ("0"+d.getDate()).slice(-2) + "-" + Math.floor(1000+Math.random()*9000);

    if (payload.action === 'RETURN_CLIENT') {
        if (payload.refDoc) {
            finalDocId = payload.refDoc.startsWith('RTN-') ? payload.refDoc : 'RTN-' + payload.refDoc;
        } else if (payload.sourceDoc) {
            let prefixMatch = payload.sourceDoc.match(/^[A-Z]+/);
            let prefix = prefixMatch ? prefixMatch[0] : 'DOC';
            finalDocId = payload.sourceDoc.replace(prefix + '-', 'RTN-');
        }
    } else if (payload.action === 'USAGE') {
        if (payload.refDoc) {
            finalDocId = payload.refDoc.startsWith('USE-') ? payload.refDoc : 'USE-' + payload.refDoc;
        } else if (payload.sourceDoc) {
            let prefixMatch = payload.sourceDoc.match(/^[A-Z]+/);
            let prefix = prefixMatch ? prefixMatch[0] : 'DOC';
            finalDocId = payload.sourceDoc.replace(prefix + '-', 'USE-');
        }
    }

    const logEntries = [];
    const requestEntries = [];
    const discrepancyEntries = [];
    const poAssignEntries = [];
    let returnedItems = [];

    // PO number for this transaction (DR encode picks it; receive carries the auto-populated value)
    let actualPoNumber = payload.poToFollow ? "" : (payload.poNumber || "");
    let poAssignStatus = payload.poToFollow ? "Pending" : "Assigned";

    if (payload.action === 'RECEIVE_DR') {
      let rData = rSheet.getDataRange().getValues();
      payload.items.forEach(itemReq => {
         let cleanCode = itemReq.code.toString().trim();
         // drId comparison fixed
         let rIdx = rData.findIndex(row => row[0].toString() === payload.drId.toString() && row[7].toString().trim().toLowerCase() === cleanCode.toLowerCase() && row[11] === 'Pending DR');
         if (rIdx !== -1) {
            rSheet.getRange(rIdx + 1, 12).setValue('Completed'); 
         }
         
         let idx = invData.findIndex((row, index) => index > 0 && row[0].toString().trim().toLowerCase() === cleanCode.toLowerCase());
         if (idx === -1) {
            let newRow = new Array(headers.length).fill(0);
            newRow[0] = cleanCode; newRow[1] = itemReq.name; newRow[2] = itemReq.uom;
            invData.push(newRow); idx = invData.length - 1;
         }
         
         let locBefore = locColIdx !== -1 ? (Number(invData[idx][locColIdx]) || 0) : 0;
         let addedQty = Number(itemReq.qty) || 0;
         
         if (locColIdx !== -1) invData[idx][locColIdx] = locBefore + addedQty;
         
         logEntries.push([ d, payload.drId, 'RECEIVE_DR', payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, itemReq.name, itemReq.uom, itemReq.wbs, itemReq.qty, locBefore, locBefore + addedQty, "Completed", "", actualPoNumber, 0, 0 ]);
      });
    }
    else {
      for (let itemReq of payload.items) {
        let cleanCode = itemReq.code.toString().trim();
        let cleanName = itemReq.name.toString().trim();
        if (!cleanCode && !cleanName) continue;
        
        let idx = invData.findIndex((row, index) => index > 0 && row[0].toString().trim().toLowerCase() === cleanCode.toLowerCase());
        
        if (idx === -1) {
          let newRow = new Array(headers.length).fill(0);
          newRow[0] = cleanCode; newRow[1] = cleanName; newRow[2] = itemReq.uom;
          invData.push(newRow); idx = invData.length - 1;
        }
        
        let locBefore = locColIdx !== -1 ? (Number(invData[idx][locColIdx]) || 0) : 0;
        let siteBefore = siteColIdx !== -1 ? (Number(invData[idx][siteColIdx]) || 0) : 0;
        let actualQty = Math.abs(Number(itemReq.qty)) || 0;
        
        let logBefore = 0; let logAfter = 0; let statusStr = "Completed";

        if (payload.action === "DR_CREATE") {
          requestEntries.push([ finalDocId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending DR', '' ]);
          logEntries.push([ d, finalDocId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, locBefore, locBefore, "Pending DR", "", actualPoNumber, itemReq.price || 0, itemReq.subtotal || 0 ]);
          poAssignEntries.push([ d, finalDocId, payload.user, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, actualPoNumber, poAssignStatus ]);
          continue;
        }
        else if (payload.action === "PURCHASE_LOG") {
          // Admin logs purchased stock directly into the warehouse location (completed immediately)
          if (locColIdx !== -1) invData[idx][locColIdx] = locBefore + actualQty;
          logEntries.push([ d, finalDocId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, locBefore, locBefore + actualQty, "Completed", "", actualPoNumber, itemReq.price || 0, itemReq.subtotal || 0 ]);
          poAssignEntries.push([ d, finalDocId, payload.user, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, actualPoNumber, poAssignStatus ]);
          continue;
        }
        else if (payload.action === "TRANSFER_WH") {
          let sourceBefore = locBefore;
          let sourceLabel = payload.location;
          
          if (sourceBefore < actualQty) throw new Error(`Insufficient stock in ${sourceLabel} for transfer. Available: ${sourceBefore}`);
          
          invData[idx][locColIdx] = locBefore - actualQty;

          let reqId = "TRN-" + Utilities.formatDate(d, "GMT+8", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
          let targetDisplay = (payload.targetSite && payload.targetSite !== "-") ? payload.targetSite : payload.targetLoc;
          requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.targetLoc, payload.targetSite, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending Receipt', '' ]);
          logEntries.push([ d, reqId, payload.action, '-', payload.user, '-', payload.siteName || '-', payload.location, cleanCode, cleanName, itemReq.uom, '', actualQty, sourceBefore, sourceBefore - actualQty, "Pending Transfer", "" ]);
          continue;
        }
        else if (payload.action === "ISSUE") {
          if (locBefore < actualQty) throw new Error(`Insufficient Location stock for ${cleanName}. Available: ${locBefore}`);
          invData[idx][locColIdx] = locBefore - actualQty;
          
          let reqId = "ISSUE-" + Utilities.formatDate(d, "GMT+8", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
          requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'In Transit', '' ]);
          logEntries.push([ d, reqId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, locBefore, locBefore - actualQty, "In Transit", "" ]);
          continue; 
        }
        else if (payload.action === "USAGE") {
          if (siteBefore < actualQty) throw new Error(`Insufficient site stock for ${cleanName}.`);
          if (siteColIdx !== -1) invData[idx][siteColIdx] = siteBefore - actualQty;
          logBefore = siteBefore; logAfter = siteBefore - actualQty;
        }
        else if (payload.action === "RETURN_CLIENT") {
          let isBySite = payload.returnType === 'BY_SITE';
          let targetBefore = locBefore;
          let returnQty = (itemReq.actualReturnQty !== null && itemReq.actualReturnQty !== "") ? Math.abs(Number(itemReq.actualReturnQty)) : 0;
          
          if (returnQty === 0) continue; 
          if (targetBefore < returnQty) throw new Error(`Insufficient physical stock in warehouse to return ${cleanName}. Available: ${targetBefore}`);
          
          if (!isBySite && actualQty < returnQty) throw new Error(`Cannot return more ${cleanName} than the document holds.`);

          invData[idx][locColIdx] = locBefore - returnQty;

          if (!isBySite && returnQty !== actualQty) {
            discrepancyEntries.push([d, finalDocId, payload.user, payload.location, payload.siteName, cleanCode, cleanName, actualQty, returnQty]);
          }

          returnedItems.push({ code: cleanCode, name: cleanName, qty: returnQty, uom: itemReq.uom });
          
          actualQty = returnQty;
          logBefore = targetBefore; 
          logAfter = targetBefore - returnQty;
        }
        else if (payload.action === "RETURN_WH") {
          if (siteBefore < actualQty) throw new Error(`Insufficient site stock for ${cleanName}.`);
          if (siteColIdx !== -1) invData[idx][siteColIdx] = siteBefore - actualQty;
          
          let reqId = "RTN-" + Utilities.formatDate(d, "GMT+8", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
          requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending Return', '' ]);
          logEntries.push([ d, reqId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, siteBefore, siteBefore - actualQty, "Pending Return", "" ]);
          continue;
        }
        
        logEntries.push([ d, finalDocId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, logBefore, logAfter, statusStr, payload.mrcNum || "" ]);
      }
    }

    const width = headers.length;
    const rectangularData = invData.map(row => {
      let newRow = [...row];
      while (newRow.length < width) newRow.push(0);
      return newRow;
    });
    
    iSheet.getRange(1, 1, rectangularData.length, width).setValues(rectangularData);

    if (payload.action === "RETURN_CLIENT" && returnedItems.length > 0) {
      try {
        let html = `
          <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 40px; color: #334155;">
            <div style="text-align: center; border-bottom: 2px solid #0f172a; padding-bottom: 20px; margin-bottom: 30px;">
              <h1 style="color: #ff5501; margin: 0; font-size: 28px; text-transform: uppercase; letter-spacing: 1px;">Return Receipt</h1>
            </div>
            <table style="width: 100%; margin-bottom: 30px; border: none;">
              <tr>
                <td style="vertical-align: top; width: 50%;">
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Document ID:</strong> ${finalDocId}</p>
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">MRC Number:</strong> ${payload.mrcNum}</p>
                </td>
                <td style="vertical-align: top; width: 50%; text-align: right;">
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Warehouse Location:</strong> ${payload.location}</p>
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Date:</strong> ${new Date().toLocaleString()}</p>
                </td>
              </tr>
            </table>
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 40px;">
              <thead>
                <tr style="background-color: #0f172a; color: #ffffff;">
                  <th style="padding: 12px 15px; text-align: left; border: 1px solid #0f172a;">Item Code</th>
                  <th style="padding: 12px 15px; text-align: left; border: 1px solid #0f172a;">Item Description</th>
                  <th style="padding: 12px 15px; text-align: right; border: 1px solid #0f172a;">Qty Returned</th>
                </tr>
              </thead>
              <tbody>`;
              
        returnedItems.forEach(ri => {
           html += `
             <tr>
               <td style="padding: 12px 15px; border: 1px solid #e2e8f0; border-top: none;">${ri.code}</td>
               <td style="padding: 12px 15px; border: 1px solid #e2e8f0; border-top: none;">${ri.name}</td>
               <td style="padding: 12px 15px; text-align: right; border: 1px solid #e2e8f0; border-top: none;"><strong>${ri.qty}</strong> ${ri.uom || 'units'}</td>
             </tr>`;
        });
        
        html += `
              </tbody>
            </table>
          </div>`;

        const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(finalDocId + ".pdf");
        const folder = DriveApp.getFolderById("1-5QL4q45Z4JvxBpB4xI6dUWHFwQ8U4cz");
        const file = folder.createFile(blob);
        const pdfUrl = file.getUrl();
        logEntries.forEach(le => { if(le[2] === 'RETURN_CLIENT') le[16] = pdfUrl; });

        const uSheet = SS.getSheetByName(SHEETS.USERS);
        const userRow = uSheet.getDataRange().getValues().find(r => r[1].toString().trim() === payload.user.toString().trim());
        if (userRow && userRow[0]) {
          GmailApp.sendEmail(userRow[0], `Return to Client Details - ${finalDocId}`, `Hello ${payload.user},\n\nPlease find attached the document details for your recent item returns from ${payload.location}.\n\nDocument ID: ${finalDocId}\n\nThank you,\nInventory System`, { attachments: [file] });
        }
      } catch (pdfErr) { console.error("PDF generation failed: " + pdfErr); }
    }

    // Normalize every log row to 20 cols: 18=PO Number (""), 19=Unit Price (0), 20=Subtotal (0)
    if (logEntries.length > 0) {
      const finalLogs = logEntries.map(le => {
        const row = le.slice(0, 20);
        while (row.length < 20) row.push(row.length === 17 ? "" : 0);
        return row;
      });
      lSheet.getRange(lSheet.getLastRow() + 1, 1, finalLogs.length, 20).setValues(finalLogs);
    }
    if (discrepancyEntries.length > 0) dSheet.getRange(dSheet.getLastRow() + 1, 1, discrepancyEntries.length, 9).setValues(discrepancyEntries);
    if (requestEntries.length > 0) rSheet.getRange(rSheet.getLastRow() + 1, 1, requestEntries.length, 13).setValues(requestEntries);

    if (poAssignEntries.length > 0) {
      let paSheet = SS.getSheetByName(SHEETS.PO_ASSIGN);
      if (!paSheet) {
        paSheet = SS.insertSheet(SHEETS.PO_ASSIGN);
        paSheet.getRange(1, 1, 1, 11).setValues([['Timestamp', 'Doc Number', 'User', 'Location', 'Site Name', 'Item Code', 'Item Description', 'UOM', 'Qty', 'PO Number', 'Status']])
          .setFontWeight("bold").setBackground("#0f172a").setFontColor("white");
        paSheet.setFrozenRows(1);
      }
      paSheet.getRange(paSheet.getLastRow() + 1, 1, poAssignEntries.length, 11).setValues(poAssignEntries);
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function processQueueAction(reqId, action, userProfile) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const rSheet = SS.getSheetByName(SHEETS.REQ);
    const iSheet = SS.getSheetByName(SHEETS.INV);
    const lSheet = SS.getSheetByName(SHEETS.LOGS);
    
    const reqData = rSheet.getDataRange().getValues();
    const rowIndex = reqData.findIndex(r => r[0] === reqId);
    if (rowIndex === -1) throw new Error("Request not found.");
    const currentData = reqData[rowIndex];
    const rowNum = rowIndex + 1;
    const reqAction = currentData[4];
    
    if (currentData[11] === 'Completed' || currentData[11] === 'Rejected') return { error: "Already processed." };

    let invData = iSheet.getDataRange().getValues();
    let headers = invData[0];
    let locColIdx = headers.indexOf(currentData[5]);
    let siteColIdx = headers.indexOf(currentData[6]);
    
    let iIdx = invData.findIndex(row => row[0].toString().trim().toLowerCase() === currentData[7].toString().trim().toLowerCase());
    if (iIdx === -1) throw new Error("Item not found in Inventory.");

    let finalQty = Number(currentData[10]) || 0;
    let newStatus = currentData[11];

    if (action === 'Reject') {
      newStatus = 'Rejected';
      
      if (reqAction === 'ISSUE' || reqAction === 'TRANSFER_WH') {
         if (locColIdx === -1) {
            locColIdx = headers.length;
            iSheet.getRange(1, locColIdx + 1).setValue(currentData[5]);
         }
         let locStock = (Number(invData[iIdx][locColIdx]) || 0);
         iSheet.getRange(iIdx + 1, locColIdx + 1).setValue(locStock + finalQty);
         lSheet.appendRow([new Date(), reqId, 'REJECTED (Reverted)', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, locStock, locStock + finalQty, 'Rejected', '']);
      } else if (reqAction === 'RETURN_WH') {
         if (siteColIdx === -1) {
            siteColIdx = headers.length;
            iSheet.getRange(1, siteColIdx + 1).setValue(currentData[6]);
         }
         let siteStock = (Number(invData[iIdx][siteColIdx]) || 0);
         iSheet.getRange(iIdx + 1, siteColIdx + 1).setValue(siteStock + finalQty); 
         lSheet.appendRow([new Date(), reqId, 'REJECTED (Reverted)', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, siteStock, siteStock + finalQty, 'Rejected', '']);
      } else {
         lSheet.appendRow([new Date(), reqId, 'REJECTED', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, 0, 0, 'Rejected', '']);
      }
      
    } else if (action === 'Confirm Receipt') {
      newStatus = 'Completed';
      if (reqAction === 'ISSUE') {
         if (siteColIdx === -1) {
            siteColIdx = headers.length;
            iSheet.getRange(1, siteColIdx + 1).setValue(currentData[6]);
         }
         let siteStock = (Number(invData[iIdx][siteColIdx]) || 0);
         iSheet.getRange(iIdx + 1, siteColIdx + 1).setValue(siteStock + finalQty);
         lSheet.appendRow([new Date(), reqId, 'TL RECEIVED', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, siteStock, siteStock + finalQty, 'Completed', '']);
      } else if (reqAction === 'TRANSFER_WH') {
         let targetName = (currentData[6] && currentData[6] !== '-') ? currentData[6] : currentData[5];
         let targetIdx = headers.indexOf(targetName);
         if (targetIdx === -1) {
            targetIdx = headers.length;
            iSheet.getRange(1, targetIdx + 1).setValue(targetName);
         }
         let stock = (Number(invData[iIdx][targetIdx]) || 0);
         iSheet.getRange(iIdx + 1, targetIdx + 1).setValue(stock + finalQty);
         lSheet.appendRow([new Date(), reqId, 'WH RECEIVED', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, stock, stock + finalQty, 'Completed', '']);
      }
    } else if (action === 'Confirm Return') {
      newStatus = 'Completed';
      if (locColIdx === -1) {
         locColIdx = headers.length;
         iSheet.getRange(1, locColIdx + 1).setValue(currentData[5]);
      }
      let locStock = (Number(invData[iIdx][locColIdx]) || 0);
      iSheet.getRange(iIdx + 1, locColIdx + 1).setValue(locStock + finalQty);
      lSheet.appendRow([new Date(), reqId, 'WH RCVD RETURN', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, locStock, locStock + finalQty, 'Completed', '']);
    }

    rSheet.getRange(rowNum, 12).setValue(newStatus); 
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { error: e.toString().replace("Error: ", "") }; } finally { lock.releaseLock(); }
}

function assignPOToDoc(docNumber, poNumber) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (!SS) throw new Error("Could not access spreadsheet.");
    if (!docNumber || !poNumber) throw new Error("Doc Number and PO Number are both required.");

    const paSheet = SS.getSheetByName(SHEETS.PO_ASSIGN);
    if (!paSheet || paSheet.getLastRow() < 2) throw new Error("PO Assignments sheet has no rows.");

    const data = paSheet.getRange(2, 1, paSheet.getLastRow() - 1, 11).getValues();
    let updated = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === docNumber.toString().trim() && data[i][10] === 'Pending') {
        paSheet.getRange(i + 2, 10).setValue(poNumber);
        paSheet.getRange(i + 2, 11).setValue('Assigned');
        updated++;
      }
    }
    if (updated === 0) throw new Error("No pending rows found for that Doc Number.");

    // Back-fill the PO Number (col 18) on the matching Logs rows so receives auto-populate it
    const lSheet = SS.getSheetByName(SHEETS.LOGS);
    if (lSheet && lSheet.getLastRow() > 1) {
      const lData = lSheet.getRange(2, 1, lSheet.getLastRow() - 1, 18).getValues();
      for (let i = 0; i < lData.length; i++) {
        if (lData[i][1] && lData[i][1].toString().trim() === docNumber.toString().trim() && !lData[i][17]) {
          lSheet.getRange(i + 2, 18).setValue(poNumber);
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, updated };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getLogoSafe() {
  try {
    const fileId = '1y-fSgltpWZthKy-SzwUpWCR9MbPWPe3o';
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return "";
  }
}

function validateUserProfile(userProfile) {
  if (!userProfile || !userProfile.email) {
    throw new Error("Authentication required.");
  }
  const uSheet = SS.getSheetByName(SHEETS.USERS);
  if (!uSheet) throw new Error("Users sheet missing.");
  const data = uSheet.getDataRange().getValues();
  const claimed = userProfile.email.toString().trim().toLowerCase();
  const row = data.find(r => r[0] && r[0].toString().trim().toLowerCase() === claimed);
  if (!row) {
    throw new Error("User account not recognized.");
  }
  return {
    email: row[0].toString().trim(),
    fullName: row[1] ? row[1].toString() : '',
    role: row[4] ? row[4].toString().toLowerCase() : 'team leader',
    locAccess: row[5] ? row[5].toString().trim() : '',
    siteAccess: row[6] ? row[6].toString().trim() : ''
  };
}