function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    { 
      name: 'Inventory', 
      head: ['Item Code', 'Item Description', 'UOM', 'Central WH', 'NCR Hub', 'Makati Site', 'Taguig Site', 'Visayas Hub', 'Cebu Site'],
      data: [
        ['ITM-001', 'Dell Latitude 5420 Laptop', 'pc', 100, 15, 10, 5, 8, 8],
        ['ITM-002', 'Logitech Wireless Mouse', 'pc', 200, 50, 30, 20, 12, 12],
        ['CBL-100', 'Cat6 Ethernet Cable (300m Box)', 'box', 50, 4, 4, 0, 0, 0]
      ]
    },
    { 
      name: 'Masterlist', 
      head: ['Client Name', 'Item Code', 'Item Description', 'UOM', 'WBS'],
      data: [
        ['Acme Corp', 'ITM-001', 'Dell Latitude 5420 Laptop', 'pc', 'WBS-991'],
        ['Acme Corp', 'ITM-002', 'Logitech Wireless Mouse', 'pc', 'WBS-991'],
        ['Globex Inc', 'CBL-100', 'Cat6 Ethernet Cable (300m Box)', 'box', 'WBS-882'],
        ['Initech', 'SVR-050', 'Cisco Catalyst 9200 Switch', 'unit', 'WBS-773']
      ]
    },
    { 
      name: 'Requests', 
      head: ['Req ID', 'Date', 'User Name', 'Role', 'Action', 'Target Location', 'Target Site', 'Item Code', 'Item Description', 'UOM', 'Qty', 'Status', 'Remarks'] 
    },
    { 
      name: 'Logs', 
      head: ['Timestamp', 'Doc Number', 'Action', 'Client Name', 'User', 'Site ID', 'Site Name', 'Warehouse Location', 'Item Code', 'Item Description', 'UOM', 'WBS', 'Qty', 'Balance Before', 'Balance After', 'Status', 'PDF Link'] 
    },
    { 
      name: 'Users', 
      head: ['Email', 'Full Name', 'Password', 'Salt', 'Role', 'Location Access', 'Site Access'],
      data: [
        ['wh@test.com', 'Alice Warehouseman', 'temp123', '', 'warehouseman', '', ''],
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
      head: ['Clients', 'Warehouse Locations', 'Site IDs', 'Site Names', 'Mapped Location'],
      data: [
        ['Acme Corp', 'NCR Hub', 'S-001', 'Makati Site', 'NCR Hub'],
        ['Globex Inc', 'Visayas Hub', 'S-002', 'Taguig Site', 'NCR Hub'],
        ['Initech', '', 'S-003', 'Cebu Site', 'Visayas Hub'],
        ['Stark Ind', '', 'S-004', 'Mandaue Site', 'Visayas Hub']
      ]
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