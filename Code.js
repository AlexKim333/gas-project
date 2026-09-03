// 상수 정의 수정 (시트 이름 확인 후 필요 시 변경)
const SHEETS = {
  STOCK: '재고시트',
  PENDING: 'PendingSheet',
  IN_LOCATIONS: '입고처목록',
  OUT_LOCATIONS: '출고처목록',
  ADMINS: '관리자명단', // 스프레드시트에서 정확한 이름으로 수정
  MANUFACTURERS: '메이커'
};

const DEFAULTS = {
  COLOR: 'SURTIDO',
  INVOICE_SUFFIX: '001'
};

// 시트 캐싱
const sheetCache = {};

function getStockData() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 최대 10초 대기
    const sheet = getSheet(SHEETS.STOCK);
    const data = sheet.getDataRange().getValues();
    const items = data.slice(1).map(row => ({
      name: String(row[0]),
      color: row[1] || DEFAULTS.COLOR,
      stockBox: row[2] || 0,
      stockIndividual: row[3] || 0,
      boxContent: row[5] || 0,
      key: `${row[0]}_${row[1] || DEFAULTS.COLOR}_${row[5] || 0}`
    }));
    console.log(`getStockData: Loaded ${items.length} items`);
    return items;
  } catch (e) {
    console.error(`getStockData error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function getManufacturers() {
  return getDropdownData(SHEETS.MANUFACTURERS, 'A');
}

function getAdminList() {
  const sheet = getSheet(SHEETS.ADMINS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.warn(`${SHEETS.ADMINS} 시트에 데이터가 없습니다 (A2부터).`);
    return [];
  }
  const admins = sheet.getRange('A2:A' + lastRow).getValues().flat().filter(item => item);
  console.log(`${SHEETS.ADMINS}에서 읽은 데이터: ${admins}`);
  return admins;
}

function registerProduct(tableData) {
  const sheet = getSheet(SHEETS.STOCK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const stockData = data.slice(1);
  const stockMap = Object.create(null);

  stockData.forEach(row => {
    const key = `${row[0]}_${row[1] || DEFAULTS.COLOR}_${row[5] || 0}`;
    stockMap[key] = { box: row[2] || 0, individual: row[3] || 0, safeStock: row[4] || 0, boxContent: row[5] || 0, initialStock: row[6] || 0 ,manufacturer: row[7] || ''};
  });

  const results = tableData.map(record => {
    const key = `${record.itemName}_${record.color || DEFAULTS.COLOR}_${record.boxContent || 0}`;
    if (stockMap[key]) {
      return { success: false, message: '기존에 같은 상품이 있습니다.', record };
    }
    const initialStock = Math.abs(record.initialStock) || 0;
    stockMap[key] = {
      box: initialStock,
      individual: 0,
      safeStock: record.safeStock || 0,
      boxContent: record.boxContent || 0,
      initialStock: initialStock,
      manufacturer: record.manufacturer || ''
    };
    return { success: true, record };
  });

  const updatedData = Object.entries(stockMap).map(([key, value]) => {
    const [name, color, boxContent] = key.split('_');
    return [name, color, value.box, value.individual, value.safeStock, parseInt(boxContent), value.initialStock, value.manufacturer];
  });
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (updatedData.length > 0) {
    sheet.getRange(2, 1, updatedData.length, headers.length).setValues(updatedData);
  }

  return results;
}

function getSheet(name) {
  if (!sheetCache[name]) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      const sheets = ss.getSheets();
      const sheetNames = sheets.map(s => s.getName());
      console.log(`시트 목록: ${sheetNames}`);
      throw new Error(`${name} 시트를 찾을 수 없습니다. 현재 시트: ${sheetNames.join(', ')}`);
    }
    sheetCache[name] = sheet;
    console.log(`시트 ${name} 성공적으로 참조됨`);
  }
  return sheetCache[name];
}

function formatDate(date) {
  return Utilities.formatDate(date, 'GMT', 'yyyy/MM/dd');
}

function getDropdownData(sheetName, column) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.warn(`${sheetName} 시트에 데이터가 없습니다 (A2부터).`);
    return [];
  }
  const data = sheet.getRange('A2:A' + lastRow).getValues().flat().filter(item => item);
  console.log(`${sheetName}에서 읽은 데이터: ${data}`);
  return data;
}

function getFilteredItemNames(searchText = '') {
  const sheet = getSheet(SHEETS.STOCK);
  const data = sheet.getDataRange().getValues();
  const items = data.slice(1).map(row => {
    const itemName = String(row[0]);
    return {
      name: itemName,
      color: row[1] || DEFAULTS.COLOR,
      stockBox: row[2] || 0,
      stockIndividual: row[3] || 0,
      safeStock: row[4] || 0,
      boxContent: row[5] || 0,
      manufacturer: row[7] || '',
      key: `${itemName}_${row[1] || DEFAULTS.COLOR}_${row[5] || 0}`
    };
  }).filter(item => !searchText || item.name.toLowerCase().includes(searchText.toLowerCase()));
  
  return items.length > 0 ? items : [{
    name: '기본품목',
    color: DEFAULTS.COLOR,
    stockBox: 0,
    stockIndividual: 0,
    safeStock: 0,
    boxContent: 0,
    manufacturer: '',
    key: `기본품목_${DEFAULTS.COLOR}_0`
  }];
}

function generateInvoiceNumber(type) {
  const sheet = getSheet(SHEETS.PENDING);
  const data = sheet.getDataRange().getValues();
  const todayStr = formatDate(new Date());
  const maxSeq = data.slice(1)
    .filter(row => row[0] && row[0].startsWith(todayStr + '-') && row[1] === type)
    .reduce((max, row) => Math.max(max, parseInt(row[0].split('-')[1]) || 0), 0);
  return String(maxSeq + 1).padStart(3, '0');
}

function getInitialInvoiceNumber() {
  return generateInvoiceNumber('입고');
}

function getInitialOutInvoiceNumber() {
  return generateInvoiceNumber('출고');
}

function processForm(tableData, type, admin) {
  const sheet = getSheet(SHEETS.PENDING);
  const stockSheet = getSheet(SHEETS.STOCK);
  const todayStr = formatDate(new Date());
  const seq = generateInvoiceNumber(type === 'in' ? '입고' : '출고');
  const invoiceNumber = `${todayStr}-${seq}`;
  
  const pendingData = tableData.map(record => {
    const manufacturer = getManufacturer(record.itemName, record.color, record.boxContent, stockSheet);
    return [
      invoiceNumber,
      type === 'in' ? '입고' : '출고',
      new Date(),
      record.itemName,
      record.color,
      Math.abs(record.boxQty) || 0,
      Math.abs(record.individualQty) || 0,
      record.boxContent || 0,
      record.location,
      admin,
      manufacturer 
    ];
  });

  const lastRow = sheet.getLastRow();
  if (pendingData.length > 0) {
    sheet.getRange(lastRow + 1, 1, pendingData.length, pendingData[0].length).setValues(pendingData);
  }
  
  updateStockSheet(tableData, type);
  return seq;
}

function getManufacturer(itemName, color, boxContent, stockSheet) {
  const data = stockSheet.getDataRange().getValues();
  const stockData = data.slice(1);
  for (let row of stockData) {
    if (row[0] === itemName && row[1] === color && row[5] === boxContent) {
      return row[7] || ''; // 재고 시트에서 제조사 정보 반환 (H열, 8번째 열)
    }
  }
  return ''; // 제조사 정보가 없는 경우 빈 문자열 반환
}

function updateStockSheet(tableData, mode) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 최대 10초 대기
    const sheet = getSheet(SHEETS.STOCK);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const stockData = data.slice(1);
    const stockMap = Object.create(null);

    stockData.forEach(row => {
      const key = `${row[0]}_${row[1] || DEFAULTS.COLOR}_${row[5] || 0}`;
      stockMap[key] = {
        box: row[2] || 0,
        individual: row[3] || 0,
        safeStock: row[4] || 0,
        boxContent: row[5] || 0,
        initialStock: row[6] || 0,
        manufacturer: row[7] || ''
      };
    });

    tableData.forEach(record => {
      const key = `${record.itemName}_${record.color || DEFAULTS.COLOR}_${record.boxContent || 0}`;
      const current = stockMap[key] || { box: 0, individual: 0, safeStock: 0, boxContent: record.boxContent || 0, initialStock: 0, manufacturer: record.manufacturer || '' };
      
      if (record.unitType === 'box') {
        const qty = Math.abs(record.boxQty) || 0;
        if (mode === 'in') {
          current.box += qty;
        } else if (mode === 'out') {
          if (current.box < qty) throw new Error('박스 재고가 부족합니다.');
          current.box -= qty;
        }
        stockMap[key] = current;
      }
    });

    tableData.forEach(record => {
      if (record.unitType === 'individual') {
        const key = `${record.itemName}_${record.color || DEFAULTS.COLOR}_${record.boxContent || 0}`;
        const current = stockMap[key] || { box: 0, individual: 0, safeStock: 0, boxContent: record.boxContent || 0, initialStock: 0 };
        const qty = Math.abs(record.individualQty) || 0;
        
        if (mode === 'in') {
          current.individual += qty;
        } else if (mode === 'out') {
          while (current.individual < qty && current.box > 0) {
            if (current.boxContent <= 0) throw new Error('박스당 내용물 개수를 입력해야 합니다.');
            current.box -= 1;
            current.individual += current.boxContent;
          }
          if (current.individual < qty) throw new Error('낱개 재고가 부족합니다. 박스 재고도 충분하지 않습니다.');
          current.individual -= qty;
        }
        stockMap[key] = current;
      }
    });

    const updatedData = Object.entries(stockMap).map(([key, value]) => {
      const [name, color, boxContent] = key.split('_');
      return [name, color, value.box, value.individual, value.safeStock, parseInt(boxContent), value.initialStock, value.manufacturer];
    });

    // 배치 쓰기로 최적화
    if (updatedData.length > 0) {
      sheet.getRange(2, 1, updatedData.length, headers.length).setValues(updatedData);
    } else {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
    }
    console.log(`updateStockSheet: Updated ${updatedData.length} items for mode ${mode}`);
  } catch (e) {
    console.error(`updateStockSheet error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function processIndividualOutStock(tableData, stockMap, sheet) {
  const headers = ['품명', '컬러', '총수량(박스)', '총수량(개)', '안전재고', '박스당낱개수량', '기초재고'];

  tableData.forEach(record => {
    if (record.unitType === 'individual' && record.individualQty < 0) {
      const key = `${record.itemName}_${record.color || DEFAULTS.COLOR}_${record.boxContent || 0}`;
      const current = stockMap[key] || { box: 0, individual: 0, safeStock: 0, boxContent: record.boxContent || 0, initialStock: 0 };
      const outQty = Math.abs(record.individualQty);
      const boxContent = record.boxContent || 0;

      while (current.individual < outQty && current.box > 0) {
        if (boxContent <= 0) throw new Error('박스당 내용물 개수를 입력해야 합니다.');
        current.box -= 1;
        current.individual += boxContent;
      }

      current.individual -= outQty;
      if (current.individual < 0) throw new Error('낱개 재고가 부족합니다. 박스 재고도 충분하지 않습니다.');
      stockMap[key] = current;
    }
  });

  const updatedData = Object.entries(stockMap).map(([key, value]) => {
    const [name, color, boxContent] = key.split('_');
    return [name, color, value.box, value.individual, value.safeStock, parseInt(boxContent), value.initialStock];
  });
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (updatedData.length > 0) {
    sheet.getRange(2, 1, updatedData.length, headers.length).setValues(updatedData);
  }
}


function searchRecords(type, invoiceNumber) {
  const sheet = getSheet(SHEETS.PENDING);
  const data = sheet.getDataRange().getValues();
  const records = data.slice(1).filter(row => row[0] === invoiceNumber && row[1] === type);
  return records.map(row => ({
    itemName: row[3],
    color: row[4],
    boxQty: row[5],
    individualQty: row[6],
    boxContent: row[7],
    location: row[8],
    admin: row[9],
    manufacturer: row[10] 
  }));
}

function updatePendingRecords(invoiceNumber, type, newRecords, admin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PENDING);
  const stockSheet = ss.getSheetByName(SHEETS.STOCK);

  if (!sheet || !stockSheet) {
    throw new Error(`필요한 시트를 찾을 수 없습니다. ${SHEETS.PENDING} 또는 ${SHEETS.STOCK} 시트가 존재하는지 확인하세요.`);
  }

  // 1. 기존 레코드 삭제 및 재고 반환
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][0] === invoiceNumber && data[i][1] === type) {
      rowsToDelete.push(i + 1);

      // 재고 반환
      const itemName = data[i][3];
      const color = data[i][4];
      const boxQty = data[i][5];
      const individualQty = data[i][6];
      const boxContent = data[i][7];

      const stockData = stockSheet.getDataRange().getValues();
      const stockRow = stockData.findIndex(row => 
        row[0] === itemName && row[1] === color && row[5] === boxContent
      );

      if (stockRow !== -1) {
        const currentBoxStock = stockData[stockRow][2];
        const currentIndividualStock = stockData[stockRow][3];

        if (type === '입고') {
          // 입고 레코드 삭제 시 재고 감소
          stockSheet.getRange(stockRow + 1, 3).setValue(currentBoxStock - boxQty);
          stockSheet.getRange(stockRow + 1, 4).setValue(currentIndividualStock - individualQty);
        } else {
          // 출고 레코드 삭제 시 재고 증가
          stockSheet.getRange(stockRow + 1, 3).setValue(currentBoxStock + boxQty);
          stockSheet.getRange(stockRow + 1, 4).setValue(currentIndividualStock + individualQty);
        }
      }
    }
  }

  // 실제 삭제
  rowsToDelete.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));

  // 2. 새 레코드 추가 및 재고 업데이트
  if (newRecords && newRecords.length > 0) {
    const newRows = newRecords.map(record => {
      const manufacturer = getManufacturer(record.itemName, record.color, record.boxContent, stockSheet);
      return [
        invoiceNumber,
        type,
        new Date(),
        record.itemName,
        record.color,
        Math.abs(record.boxQty),
        Math.abs(record.individualQty),
        record.boxContent || 0,
        record.location,
        admin,
        manufacturer
      ];
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);

      // 재고 업데이트
      newRecords.forEach(record => {
        const stockData = stockSheet.getDataRange().getValues();
        const stockRow = stockData.findIndex(row => 
          row[0] === record.itemName && row[1] === record.color && row[5] === record.boxContent
        );

        if (stockRow !== -1) {
          const currentBoxStock = stockData[stockRow][2];
          const currentIndividualStock = stockData[stockRow][3];

          if (type === '입고') {
            stockSheet.getRange(stockRow + 1, 3).setValue(currentBoxStock + Math.abs(record.boxQty));
            stockSheet.getRange(stockRow + 1, 4).setValue(currentIndividualStock + Math.abs(record.individualQty));
          } else {
            const newBoxStock = currentBoxStock - Math.abs(record.boxQty);
            const newIndividualStock = currentIndividualStock - Math.abs(record.individualQty);
            if (newBoxStock < 0) throw new Error('박스 재고가 부족합니다.');
            if (newIndividualStock < 0) throw new Error('낱개 재고가 부족합니다.');
            stockSheet.getRange(stockRow + 1, 3).setValue(newBoxStock);
            stockSheet.getRange(stockRow + 1, 4).setValue(newIndividualStock);
          }
        } else {
          stockSheet.appendRow([
            record.itemName,
            record.color,
            type === '입고' ? Math.abs(record.boxQty) : -Math.abs(record.boxQty),
            type === '입고' ? Math.abs(record.individualQty) : -Math.abs(record.individualQty),
            0,
            record.boxContent || 0,
            0
          ]);
        }
      });
    }
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('창고 관리')
    .addItem('입고입력', 'showAppIn')
    .addItem('출고입력', 'showAppOut')
    .addItem('상품등록', 'showAppProduct')
    .addItem('검색수정', 'showSearchModify')
    .addToUi();
}

function showSearchModify() {
  const html = HtmlService.createHtmlOutputFromFile('SearchModify')
    .setWidth(1075)
    .setHeight(851);
  SpreadsheetApp.getUi().showModalDialog(html, '창고 관리 - 검색수정');
}

function showAppIn() {
  showApp('in');
}

function showAppOut() {
  showApp('out');
}

function showAppProduct() {
  showApp('product');
}

function showApp(type, pendingData = null) {
  const template = HtmlService.createTemplateFromFile('WarehouseApp');
  template.currentType = type;
  template.pendingRecord = pendingData ? JSON.stringify(pendingData) : 'null';
  const html = template.evaluate().setWidth(1075).setHeight(851);
  SpreadsheetApp.getUi().showModalDialog(html, '창고 관리');
}

function getInLocations() {
  return getDropdownData(SHEETS.IN_LOCATIONS, '입고처');
}

function getOutLocations() {
  return getDropdownData(SHEETS.OUT_LOCATIONS, '출고처');
}

function processInForm(tableData, admin) {
  return processForm(tableData, 'in', admin);
}

function processOutForm(tableData, admin) {
  return processForm(tableData, 'out', admin);
}

function getMaxSequentialNumber(date, type) {
  const formattedDate = date.replace(/-/g, '/');
  const sheet = getSheet(SHEETS.PENDING);
  const data = sheet.getDataRange().getValues();
  const invoiceNumbers = data.slice(1)
    .filter(row => row[1] === type && row[0].startsWith(formattedDate + '-'))
    .map(row => row[0]);
  if (invoiceNumbers.length === 0) return 0;
  const seqNumbers = invoiceNumbers.map(inv => parseInt(inv.split('-')[1]));
  return Math.max(...seqNumbers);
}

function checkItemRegistration(itemName, color, boxContent) {
  const sheet = getSheet(SHEETS.STOCK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const stockData = data.slice(1);
  const stockMap = Object.create(null);

  stockData.forEach(row => {
    const key = `${row[0]}_${row[1] || DEFAULTS.COLOR}_${row[5] || 0}`;
    stockMap[key] = { box: row[2] || 0, individual: row[3] || 0, safeStock: row[4] || 0, boxContent: row[5] || 0, initialStock: row[6] || 0 };
  });

  const key = `${itemName}_${color || DEFAULTS.COLOR}_${boxContent || 0}`;
  return stockMap[key] !== undefined;
}

function checkStockAvailability(itemName, color, boxContent, qty) {
  const sheet = getSheet(SHEETS.STOCK); // '재고시트' 가져오기
  const data = sheet.getDataRange().getValues();
  const stockData = data.slice(1); // 헤더 제외
  let boxStock = 0;

  stockData.forEach(row => {
    if (row[0] === itemName && (row[1] || DEFAULTS.COLOR) === color && (row[5] || 0) === boxContent) {
      boxStock = row[2] || 0; // 박스 재고 (3번째 열)
    }
  });

  return { boxStock: boxStock, isSufficient: boxStock >= qty };
}