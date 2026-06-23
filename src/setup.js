function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    { 
      name: 'SKU Masterlist',
      head: ['Item Code', 'Item Description', 'UOM', 'NCR Hub', 'Makati Site', 'Taguig Site', 'Visayas Hub', 'Cebu Site'],
      data: [
        ['ITM-001', 'Dell Latitude 5420 Laptop', 'pc', 70, 30, 0, 80, 20],
        ['ITM-002', 'Logitech Wireless Mouse', 'pc', 100, 50, 50, 60, 40],
        ['CBL-100', 'Cat6 Ethernet Cable (300m Box)', 'box', 50, 0, 0, 10, 0]
      ]
    },
    {
      name: 'Requests',
      head: ['Req ID', 'Date', 'User Name', 'Role', 'Action', 'Target Location', 'Target Site', 'Item Code', 'Item Description', 'UOM', 'Qty', 'Status', 'Remarks', 'User Email']
    },
    {
      name: 'Logs',
      head: ['Timestamp', 'Doc Number', 'Action', 'Client Name', 'User', 'Site ID', 'Site Name', 'Warehouse Location', 'Item Code', 'Item Description', 'UOM', 'WBS', 'Qty', 'Balance Before', 'Balance After', 'Status', 'PDF Link', 'PO Number', 'Unit Price', 'Subtotal', 'Destination Location', 'Destination Site']
    },
    { 
      name: 'Users', 
      head: ['Email', 'Full Name', 'Password', 'Salt', 'Role', 'Location Access', 'Site Access'],
      data: [
        ['admin@test.com', 'Admin User', 'temp123', '', 'admin', '', ''],
        ['wh@test.com', 'Alice Warehouseman', 'temp123', '', 'warehouseman', 'NCR Hub', ''],
        ['tl1@test.com', 'Bob TeamLeader', 'temp123', '', 'team leader', 'NCR Hub', 'Makati Site'],
        ['tl2@test.com', 'Charlie TeamLeader', 'temp123', '', 'team leader', 'Visayas Hub', 'Cebu Site, Mandaue Site']
      ]
    },
    { 
      name: 'Discrepancies', 
      head: ['Timestamp', 'Doc ID', 'User', 'Warehouse Location', 'Site Name', 'Item Code', 'Item Description', 'System Qty', 'Actual Returned Qty'] 
    },
    { 
      name: 'Dropdowns', 
      head: ['Clients', 'Warehouse Locations', 'Site IDs', 'Site Names', 'Mapped Location', 'WBS'],
      data: [
        ['Acme Corp', 'NCR Hub', 'S-001', 'Makati Site', 'NCR Hub', 'WBS-991'],
        ['Globex Inc', 'Visayas Hub', 'S-002', 'Taguig Site', 'NCR Hub', 'WBS-882'],
        ['Initech', '', 'S-003', 'Cebu Site', 'Visayas Hub', 'WBS-773'],
        ['Stark Ind', '', 'S-004', 'Mandaue Site', 'Visayas Hub', 'WBS-774']
      ]
    },
    {
      name: 'PO Database',
      head: ['PO Number'],
      data: [
        ['PO-10001'],
        ['PO-10002'],
        ['PO-10003'],
        ['PO-10004']
      ]
    },
    {
      name: 'PO Assignments',
      head: ['Timestamp', 'Doc Number', 'User', 'Location', 'Site Name', 'Item Code', 'Item Description', 'UOM', 'Qty', 'PO Number', 'Status']
    },
    {
      name: 'Notifications',
      head: ['Notif ID', 'Timestamp', 'Recipient Email', 'Recipient Name', 'Recipient Role', 'Sender Email', 'Sender Name', 'Sender Role', 'Action', 'Related Req ID', 'Message', 'Read', 'Read At', 'Email Status']
    }
  ];

  sheets.forEach(sh => {
    let sheet = ss.getSheetByName(sh.name) || ss.insertSheet(sh.name);
    const currentCols = sheet.getLastColumn();
    const currentRows = sheet.getLastRow();

    if (currentCols === 0 || currentCols < sh.head.length) {
      sheet.getRange(1, 1, 1, sh.head.length).setValues([sh.head])
           .setFontWeight("bold").setBackground("#0f172a").setFontColor("white");
    }
    
    if(sheet.getFrozenRows() === 0) sheet.setFrozenRows(1);

    if (currentRows <= 1 && sh.data && sh.data.length > 0) {
      sheet.getRange(2, 1, sh.data.length, sh.head.length).setValues(sh.data);
    }
  });
}