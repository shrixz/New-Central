const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  INV: 'Inventory',
  LOGS: 'Logs',
  USERS: 'Users',
  REQ: 'Requests',
  DISCREP: 'Discrepancies',
  MASTER: 'Masterlist',
  DROPS: 'Dropdowns'
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
    const mSheet = SS.getSheetByName(SHEETS.MASTER);
    const uSheet = SS.getSheetByName(SHEETS.USERS);
    const rSheet = SS.getSheetByName(SHEETS.REQ);
    const dSheet = SS.getSheetByName(SHEETS.DROPS) || SS.insertSheet(SHEETS.DROPS);
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
          central: Number(invData[i][3]) || 0,
          balances: {}
        };
        for (let j = 4; j < headers.length; j++) {
           if (headers[j] && headers[j].toString().trim() !== "") {
              item.balances[headers[j]] = Number(invData[i][j]) || 0;
           }
        }
        inventory.push(item);
      }
    }

    let masterlist = [];
    if (mSheet && mSheet.getLastRow() > 1) {
      const mData = mSheet.getDataRange().getValues();
      for (let i = 1; i < mData.length; i++) {
         if (mData[i][1]) {
           masterlist.push({
             client: mData[i][0]?.toString() || '',
             code: mData[i][1]?.toString() || '',
             name: mData[i][2]?.toString() || '',
             uom: mData[i][3]?.toString() || '',
             wbs: mData[i][4]?.toString() || ''
           });
         }
      }
    }

    let dropdowns = { clients: [], locations: [], siteIds: [], sites: [], siteMap: {} };
    if (dSheet && dSheet.getLastRow() > 1) {
      const dData = dSheet.getDataRange().getValues();
      for (let i = 1; i < dData.length; i++) {
         if (dData[i][0]) dropdowns.clients.push(dData[i][0].toString());
         if (dData[i][1]) dropdowns.locations.push(dData[i][1].toString());
         if (dData[i][2]) dropdowns.siteIds.push(dData[i][2].toString());
         if (dData[i][3]) {
            let sName = dData[i][3].toString();
            dropdowns.sites.push(sName);
            let mappedLoc = dData[i][4] ? dData[i][4].toString() : "";
            if (mappedLoc) dropdowns.siteMap[sName] = mappedLoc;
         }
      }
    }

    let pending = [];
    if (rSheet && rSheet.getLastRow() > 1 && userProfile) {
      const reqData = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 13).getValues();
      reqData.forEach(r => {
        let status = r[11];
        let reqAction = r[4];
        let reqLoc = r[5];
        let reqSite = r[6];
        let isMyTurn = false;
        
        if (userProfile.role === 'warehouseman') {
           if (status === 'Pending Return' && reqAction === 'RETURN_CENTRAL') isMyTurn = true;
        }
        else if (userProfile.role === 'team leader') {
           // TL confirms items explicitly issued to their location/site
           if (status === 'In Transit' && reqAction === 'ISSUE') {
              let locMatch = (!userProfile.locAccess || userProfile.locAccess.includes(reqLoc));
              let siteMatch = (!userProfile.siteAccess || userProfile.siteAccess.includes(reqSite));
              if (locMatch && siteMatch) isMyTurn = true;
           }
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
    if (lSheet && lSheet.getLastRow() > 1) {
      const logValues = lSheet.getRange(2, 1, lSheet.getLastRow() - 1, 17).getValues();
      logs = logValues.reverse().map(l => {
        return [
          l[0].toISOString ? l[0].toISOString() : l[0], 
          l[1]||'', l[2]||'', l[3]||'', l[4]||'', l[5]||'', l[6]||'', l[7]||'', l[8]||'', l[9]||'', l[10]||'', 
          l[11]||'', l[12]||0, l[13]||0, l[14]||0, l[15]||'Completed', l[16]||''
        ];
      });
    }
    
    return { inventory, logs, masterlist, dropdowns, pending };
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
    
    if (invData.length === 0) invData = [['Item Code', 'Item Description', 'UOM', 'Central WH']];
    let headers = invData[0];
    const centralColIdx = 3;
    
    const isCentralDirect = (payload.role === 'warehouseman' && (payload.action === "IN" || payload.action === "RETURN_CLIENT"));
    
    let locColIdx = -1; let siteColIdx = -1;
    if (!isCentralDirect) {
      if (payload.location && payload.location !== "-" && payload.location !== "Central WH") {
        locColIdx = headers.indexOf(payload.location);
        if (locColIdx === -1) { headers.push(payload.location); locColIdx = headers.length - 1; }
      }
      if (payload.siteName && payload.siteName !== "-") {
        siteColIdx = headers.indexOf(payload.siteName);
        if (siteColIdx === -1) { headers.push(payload.siteName); siteColIdx = headers.length - 1; }
      }
    }
    
    const d = new Date();
    const finalDocId = payload.refDoc ? payload.refDoc : "DOC-" + d.getFullYear() + ("0"+(d.getMonth()+1)).slice(-2) + ("0"+d.getDate()).slice(-2) + "-" + Math.floor(1000+Math.random()*9000);

    const logEntries = [];
    const requestEntries = [];
    const discrepancyEntries = [];
    let returnedItems = [];

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
      
      let centralBefore = Number(invData[idx][centralColIdx]) || 0;
      let locBefore = locColIdx !== -1 ? (Number(invData[idx][locColIdx]) || 0) : 0;
      let siteBefore = siteColIdx !== -1 ? (Number(invData[idx][siteColIdx]) || 0) : 0;
      let actualQty = Math.abs(itemReq.qty);
      
      let logBefore = 0; let logAfter = 0; let statusStr = "Completed";

      if (payload.action === "IN") {
        invData[idx][centralColIdx] = centralBefore + actualQty;
        logBefore = centralBefore; logAfter = centralBefore + actualQty;
      } 
      else if (payload.action === "ISSUE") {
        if (centralBefore < actualQty) throw new Error(`Insufficient Central WH stock for ${cleanName}. Available: ${centralBefore}`);
        invData[idx][centralColIdx] = centralBefore - actualQty;
        
        let reqId = "ISSUE-" + Utilities.formatDate(d, "GMT+8", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
        requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'In Transit', '' ]);
        logEntries.push([ d, reqId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, centralBefore, centralBefore - actualQty, "In Transit", "" ]);
        continue; 
      }
      else if (payload.action === "USAGE") {
        let targetBefore = payload.siteName ? siteBefore : locBefore;
        if (targetBefore < actualQty) throw new Error(`Insufficient site stock for ${cleanName}.`);
        invData[idx][locColIdx] = locBefore - actualQty;
        if (siteColIdx !== -1) invData[idx][siteColIdx] = siteBefore - actualQty;
        logBefore = targetBefore; logAfter = targetBefore - actualQty;
      }
      else if (payload.action === "RETURN_CLIENT") {
        let targetBefore = isCentralDirect ? centralBefore : (payload.siteName && payload.siteName !== "-" ? siteBefore : locBefore);
        let actualReturnQty = (itemReq.actualReturnQty !== null && itemReq.actualReturnQty !== "") ? Math.abs(Number(itemReq.actualReturnQty)) : targetBefore;
        
        if (targetBefore < actualQty) throw new Error(`Insufficient stock to return ${cleanName}.`);
        
        if (isCentralDirect) {
          invData[idx][centralColIdx] = centralBefore - targetBefore;
        } else {
          invData[idx][locColIdx] = locBefore - targetBefore;
          if (siteColIdx !== -1) invData[idx][siteColIdx] = siteBefore - targetBefore;
        }

        if (actualReturnQty !== targetBefore) {
          discrepancyEntries.push([d, finalDocId, payload.user, payload.location, payload.siteName, cleanCode, cleanName, targetBefore, actualReturnQty]);
        }
        if (actualReturnQty > 0 || targetBefore > 0) {
          returnedItems.push({ code: cleanCode, name: cleanName, qty: actualReturnQty, uom: itemReq.uom });
        }
        
        actualQty = targetBefore;
        logBefore = targetBefore; logAfter = 0;
      }
      else if (payload.action === "RETURN_CENTRAL") {
        let targetBefore = payload.siteName ? siteBefore : locBefore;
        if (targetBefore < actualQty) throw new Error(`Insufficient site stock for ${cleanName}.`);
        invData[idx][locColIdx] = locBefore - actualQty;
        if (siteColIdx !== -1) invData[idx][siteColIdx] = siteBefore - actualQty;
        
        let reqId = "RTN-" + Utilities.formatDate(d, "GMT+8", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
        requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending Return', '' ]);
        logEntries.push([ d, reqId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, targetBefore, targetBefore - actualQty, "Pending Return", "" ]);
        continue;
      }
      
      logEntries.push([ d, finalDocId, payload.action, payload.client, payload.user, payload.siteId, payload.siteName, payload.location, cleanCode, cleanName, itemReq.uom, itemReq.wbs, actualQty, logBefore, logAfter, statusStr, "" ]);
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
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Date:</strong> ${new Date().toLocaleString()}</p>
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Client:</strong> ${payload.client || 'N/A'}</p>
                </td>
                <td style="vertical-align: top; width: 50%; text-align: right;">
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Warehouse Location:</strong> ${payload.location}</p>
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Site ID:</strong> ${payload.siteId || 'N/A'}</p>
                  <p style="margin: 4px 0;"><strong style="color: #0f172a;">Site Name:</strong> ${payload.siteName || 'N/A'}</p>
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
        logEntries.forEach(le => { le[16] = pdfUrl; });

        const uSheet = SS.getSheetByName(SHEETS.USERS);
        const userRow = uSheet.getDataRange().getValues().find(r => r[1].toString().trim() === payload.user.toString().trim());
        if (userRow && userRow[0]) {
          GmailApp.sendEmail(userRow[0], `Return to Client Details - ${finalDocId}`, `Hello ${payload.user},\n\nPlease find attached the document details for your recent item returns from ${payload.location}.\n\nDocument ID: ${finalDocId}\n\nThank you,\nInventory System`, { attachments: [file] });
        }
      } catch (pdfErr) { console.error("PDF generation failed: " + pdfErr); }
    }

    if (logEntries.length > 0) lSheet.getRange(lSheet.getLastRow() + 1, 1, logEntries.length, 17).setValues(logEntries);
    if (discrepancyEntries.length > 0) dSheet.getRange(dSheet.getLastRow() + 1, 1, discrepancyEntries.length, 9).setValues(discrepancyEntries);
    if (requestEntries.length > 0) rSheet.getRange(rSheet.getLastRow() + 1, 1, requestEntries.length, 13).setValues(requestEntries);
    
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
    
    if (currentData[11] === 'Completed' || currentData[11] === 'Rejected') return { error: "Already processed." };

    let invData = iSheet.getDataRange().getValues();
    let headers = invData[0];
    const centralColIdx = 3;
    let locColIdx = headers.indexOf(currentData[5]);
    let siteColIdx = headers.indexOf(currentData[6]);
    
    let iIdx = invData.findIndex(row => row[0].toString().trim().toLowerCase() === currentData[7].toString().trim().toLowerCase());
    if (iIdx === -1) throw new Error("Item not found in Inventory.");

    let finalQty = currentData[10];
    let newStatus = currentData[11];

    if (action === 'Confirm Receipt') {
      newStatus = 'Completed';
      let locStock = (locColIdx !== -1) ? (Number(invData[iIdx][locColIdx]) || 0) : 0;
      let siteStock = (siteColIdx !== -1) ? (Number(invData[iIdx][siteColIdx]) || 0) : 0;
      if (locColIdx !== -1) iSheet.getRange(iIdx + 1, locColIdx + 1).setValue(locStock + finalQty);
      if (siteColIdx !== -1) iSheet.getRange(iIdx + 1, siteColIdx + 1).setValue(siteStock + finalQty);
      lSheet.appendRow([new Date(), reqId, 'TL RECEIVED', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, siteStock, siteStock + finalQty, 'Completed', '']);
    } else if (action === 'Confirm Return') {
      newStatus = 'Completed';
      let centralStock = Number(invData[iIdx][centralColIdx]) || 0;
      iSheet.getRange(iIdx + 1, centralColIdx + 1).setValue(centralStock + finalQty);
      lSheet.appendRow([new Date(), reqId, 'WH RCVD RETURN', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, centralStock, centralStock + finalQty, 'Completed', '']);
    }

    rSheet.getRange(rowNum, 12).setValue(newStatus); 
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { error: e.toString().replace("Error: ", "") }; } finally { lock.releaseLock(); }
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