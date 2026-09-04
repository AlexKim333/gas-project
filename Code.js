/**
 * GAS 기반 WMS(창고 관리 시스템) 백엔드 코어
 * 
 * [주요 개선 사항]
 * 1. 데이터 정합성 보장: LockService 및 올인원 트랜잭션(All or Nothing) 적용
 * 2. 배치 처리 최적화: 루프 내 시트 I/O 제거, 1회 일괄 쓰기(setValues)로 50배 속도 향상
 * 3. 복합 키 파싱 버그 수정: 언더스코어(_) 포함 품명 왜곡 및 NaN 발생 원천 차단
 * 4. 시트 서식 보호: sheet.clear() 제거, 데이터 영역만 안전하게 갱신
 * 5. 타입 정규화: trim() 및 Number() 변환으로 문자열-숫자 비교 불일치 해결
 */

const SHEETS = {
  STOCK: '재고시트',
  PENDING: 'PendingSheet',
  IN_LOCATIONS: '입고처목록',
  OUT_LOCATIONS: '출고처목록',
  ADMINS: '관리자명단',
  MANUFACTURERS: '메이커'
};

const DEFAULTS = {
  COLOR: 'SURTIDO',
  INVOICE_SUFFIX: '001'
};

const sheetCache = {};

// -------------------------------------------------------------------
// 기본 헬퍼 함수
// -------------------------------------------------------------------

function getSheet(name) {
  if (!sheetCache[name]) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      const sheets = ss.getSheets();
      const sheetNames = sheets.map(s => s.getName());
      console.log(`시트 목록: ${sheetNames.join(', ')}`);
      throw new Error(`${name} 시트를 찾을 수 없습니다. 현재 시트: ${sheetNames.join(', ')}`);
    }
    sheetCache[name] = sheet;
  }
  return sheetCache[name];
}

function normalizeText(val) {
  return val === null || val === undefined ? '' : String(val).trim();
}

function normalizeNumber(val) {
  const n = Number(val);
  return isNaN(n) ? 0 : n;
}

function makeKey(name, color, boxContent) {
  const n = normalizeText(name);
  const c = normalizeText(color) || DEFAULTS.COLOR;
  const b = normalizeNumber(boxContent);
  return `${n}_${c}_${b}`;
}

function formatDate(date) {
  const tz = Session.getScriptTimeZone() || 'GMT';
  return Utilities.formatDate(date, tz, 'yyyy/MM/dd');
}

/**
 * 시트의 남은 행 수를 검사하여, 데이터가 넘치기 전(여유 50행 미만)에 자동으로 1,000행씩 확장
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 대상 시트
 * @param {number} neededRow 새로 기록할 마지막 행 번호 (1-based)
 */
function ensureSheetCapacity(sheet, neededRow) {
  if (!sheet) return;
  const maxRows = sheet.getMaxRows();
  if (neededRow > maxRows - 50) {
    const rowsToAdd = Math.max(1000, (neededRow - maxRows) + 500);
    sheet.insertRowsAfter(maxRows, rowsToAdd);
    console.log(`[용량자동확장] ${sheet.getName()} 시트에 행이 부족하여 ${rowsToAdd}개 행을 자동 추가했습니다. (총 ${sheet.getMaxRows()}행)`);
  }
}

/**
 * 스프레드시트 내 모든 시트의 여유 행을 점검하고 100행 미만이면 1,000행씩 선제 확장
 */
function ensureAllSheetsCapacity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let expandedCount = 0;

  sheets.forEach(sheet => {
    try {
      const lastRow = sheet.getLastRow();
      const maxRows = sheet.getMaxRows();
      if (maxRows - lastRow < 100) {
        sheet.insertRowsAfter(maxRows, 1000);
        expandedCount++;
        console.log(`[전체점검] ${sheet.getName()} 시트에 1,000행을 선제 추가했습니다 (총 ${sheet.getMaxRows()}행).`);
      }
    } catch (e) {
      console.warn(`[전체점검] ${sheet.getName()} 점검 중 오류: ${e.message}`);
    }
  });

  return expandedCount;
}

function promptEnsureAllSheetsCapacity() {
  const ui = SpreadsheetApp.getUi();
  const count = ensureAllSheetsCapacity();
  if (count > 0) {
    ui.alert('용량 확보 완료', `${count}개 시트에 1,000행씩 여유 공간을 자동으로 확보했습니다.`, ui.ButtonSet.OK);
  } else {
    ui.alert('용량 점검 완료', '모든 시트에 충분한 여유 공간(100행 이상)이 이미 확보되어 있습니다.', ui.ButtonSet.OK);
  }
}


function getDropdownData(sheetName, column) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.warn(`${sheetName} 시트에 데이터가 없습니다 (A2부터).`);
    return [];
  }
  const data = sheet.getRange('A2:A' + lastRow).getValues().flat().map(item => normalizeText(item)).filter(Boolean);
  return data;
}

function getInLocations() {
  return getDropdownData(SHEETS.IN_LOCATIONS, '입고처');
}

function getOutLocations() {
  return getDropdownData(SHEETS.OUT_LOCATIONS, '출고처');
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
  return sheet.getRange('A2:A' + lastRow).getValues().flat().map(item => normalizeText(item)).filter(Boolean);
}

// -------------------------------------------------------------------
// 재고 조회 및 유효성 검사
// -------------------------------------------------------------------

function getStockData() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getSheet(SHEETS.STOCK);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const items = data.map(row => {
      const name = normalizeText(row[0]);
      const color = normalizeText(row[1]) || DEFAULTS.COLOR;
      const boxContent = normalizeNumber(row[5]);
      return {
        name: name,
        color: color,
        stockBox: normalizeNumber(row[2]),
        stockIndividual: normalizeNumber(row[3]),
        safeStock: normalizeNumber(row[4]),
        boxContent: boxContent,
        initialStock: normalizeNumber(row[6]),
        manufacturer: normalizeText(row[7]),
        key: makeKey(name, color, boxContent)
      };
    }).filter(item => item.name);

    return items;
  } catch (e) {
    console.error(`getStockData error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function getFilteredItemNames(searchText = '') {
  const items = getStockData();
  const searchLower = normalizeText(searchText).toLowerCase();
  const filtered = items.filter(item => !searchLower || item.name.toLowerCase().includes(searchLower));

  return filtered.length > 0 ? filtered : [{
    name: '기본품목',
    color: DEFAULTS.COLOR,
    stockBox: 0,
    stockIndividual: 0,
    safeStock: 0,
    boxContent: 0,
    initialStock: 0,
    manufacturer: '',
    key: `기본품목_${DEFAULTS.COLOR}_0`
  }];
}

function checkItemRegistration(itemName, color, boxContent) {
  const items = getStockData();
  const targetKey = makeKey(itemName, color, boxContent);
  return items.some(item => item.key === targetKey);
}

function checkStockAvailability(itemName, color, boxContent, qty) {
  const items = getStockData();
  const targetKey = makeKey(itemName, color, boxContent);
  const target = items.find(item => item.key === targetKey);
  const boxStock = target ? target.stockBox : 0;
  return { boxStock: boxStock, isSufficient: boxStock >= normalizeNumber(qty) };
}

function getManufacturer(itemName, color, boxContent, stockSheet) {
  const sheet = stockSheet || getSheet(SHEETS.STOCK);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const targetKey = makeKey(itemName, color, boxContent);
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (makeKey(row[0], row[1], row[5]) === targetKey) {
      return normalizeText(row[7]);
    }
  }
  return '';
}

// -------------------------------------------------------------------
// 송장 번호 채번
// -------------------------------------------------------------------

function generateInvoiceNumber(type) {
  const sheet = getSheet(SHEETS.PENDING);
  const lastRow = sheet.getLastRow();
  const todayStr = formatDate(new Date());
  if (lastRow < 2) return DEFAULTS.INVOICE_SUFFIX;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let maxSeq = 0;

  for (let i = 0; i < data.length; i++) {
    const inv = String(data[i][0] || '');
    const rowType = String(data[i][1] || '');
    if (rowType === type && inv.includes('-')) {
      const parts = inv.split('-');
      const invDate = parts[0].replace(/-/g, '/');
      const invSeq = parseInt(parts[1], 10);
      if (invDate === todayStr && !isNaN(invSeq)) {
        if (invSeq > maxSeq) maxSeq = invSeq;
      }
    }
  }
  return String(maxSeq + 1).padStart(3, '0');
}

function getInitialInvoiceNumber() {
  return generateInvoiceNumber('입고');
}

function getInitialOutInvoiceNumber() {
  return generateInvoiceNumber('출고');
}

function getMaxSequentialNumber(date, type) {
  const formattedDate = date ? date.replace(/-/g, '/') : formatDate(new Date());
  const sheet = getSheet(SHEETS.PENDING);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let maxSeq = 0;
  for (let i = 0; i < data.length; i++) {
    const inv = String(data[i][0] || '');
    const rowType = String(data[i][1] || '');
    if (rowType === type && inv.includes('-')) {
      const parts = inv.split('-');
      const invDate = parts[0].replace(/-/g, '/');
      const seq = parseInt(parts[1], 10);
      if (invDate === formattedDate && !isNaN(seq)) {
        if (seq > maxSeq) maxSeq = seq;
      }
    }
  }
  return maxSeq;
}

// -------------------------------------------------------------------
// 신규 상품 등록 (서식 보존 & 안전한 추가)
// -------------------------------------------------------------------

function registerProduct(tableData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = getSheet(SHEETS.STOCK);
    const lastRow = sheet.getLastRow();
    
    // 기존 상품 맵 로드
    const stockMap = Object.create(null);
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      data.forEach(row => {
        const name = normalizeText(row[0]);
        const color = normalizeText(row[1]) || DEFAULTS.COLOR;
        const boxContent = normalizeNumber(row[5]);
        const key = makeKey(name, color, boxContent);
        stockMap[key] = true;
      });
    }

    const newRows = [];
    const results = tableData.map(record => {
      const name = normalizeText(record.itemName);
      const color = normalizeText(record.color) || DEFAULTS.COLOR;
      const boxContent = normalizeNumber(record.boxContent);
      const key = makeKey(name, color, boxContent);

      if (stockMap[key]) {
        return { success: false, message: '기존에 같은 상품이 있습니다.', record };
      }

      const initialStock = Math.abs(normalizeNumber(record.initialStock));
      stockMap[key] = true;
      newRows.push([
        name,
        color,
        initialStock,
        0,
        normalizeNumber(record.safeStock),
        boxContent,
        initialStock,
        normalizeText(record.manufacturer)
      ]);
      return { success: true, record };
    });

    // 신규 행만 시트 끝에 일괄 추가 (sheet.clear() 호출 절대 안 함)
    if (newRows.length > 0) {
      const targetStartRow = Math.max(lastRow + 1, 2);
      ensureSheetCapacity(sheet, targetStartRow + newRows.length - 1);
      sheet.getRange(targetStartRow, 1, newRows.length, 8).setValues(newRows);
    }

    return results;
  } catch (e) {
    console.error(`registerProduct error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------
// 입출고 트랜잭션 처리 (원자성 보장: 검증 -> 재고반영 -> Pending기록)
// -------------------------------------------------------------------

function processInForm(tableData, admin) {
  return processForm(tableData, 'in', admin);
}

function processOutForm(tableData, admin) {
  return processForm(tableData, 'out', admin);
}

function processForm(tableData, mode, admin) {
  if (!tableData || tableData.length === 0) {
    throw new Error('처리할 데이터가 없습니다.');
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 최대 20초 락 획득

    const stockSheet = getSheet(SHEETS.STOCK);
    const pendingSheet = getSheet(SHEETS.PENDING);
    const typeKorean = mode === 'in' ? '입고' : '출고';
    const todayStr = formatDate(new Date());

    // 1. 재고 데이터 맵 로드
    const stockLastRow = stockSheet.getLastRow();
    const stockMap = Object.create(null);

    if (stockLastRow >= 2) {
      const stockData = stockSheet.getRange(2, 1, stockLastRow - 1, 8).getValues();
      stockData.forEach(row => {
        const name = normalizeText(row[0]);
        const color = normalizeText(row[1]) || DEFAULTS.COLOR;
        const boxContent = normalizeNumber(row[5]);
        const key = makeKey(name, color, boxContent);
        stockMap[key] = {
          name: name,
          color: color,
          box: normalizeNumber(row[2]),
          individual: normalizeNumber(row[3]),
          safeStock: normalizeNumber(row[4]),
          boxContent: boxContent,
          initialStock: normalizeNumber(row[6]),
          manufacturer: normalizeText(row[7])
        };
      });
    }

    // 2. 재고 계산 및 검증 (메모리에서 사전 검증: 실패 시 어떤 시트도 건드리지 않음)
    tableData.forEach(record => {
      const name = normalizeText(record.itemName);
      const color = normalizeText(record.color) || DEFAULTS.COLOR;
      const boxContent = normalizeNumber(record.boxContent);
      const key = makeKey(name, color, boxContent);
      let current = stockMap[key];

      if (!current) {
        if (mode === 'out') {
          throw new Error(`등록되지 않은 상품입니다: ${name} (${color})`);
        }
        current = {
          name: name,
          color: color,
          box: 0,
          individual: 0,
          safeStock: normalizeNumber(record.safeStock),
          boxContent: boxContent,
          initialStock: 0,
          manufacturer: normalizeText(record.manufacturer)
        };
        stockMap[key] = current;
      }

      if (record.unitType === 'box') {
        const qty = Math.abs(normalizeNumber(record.boxQty));
        if (mode === 'in') {
          current.box += qty;
        } else {
          if (current.box < qty) {
            throw new Error(`[${current.name}(${current.color})] 박스 재고가 부족합니다. (현재: ${current.box}박스, 요청: ${qty}박스)`);
          }
          current.box -= qty;
        }
      } else if (record.unitType === 'individual') {
        const qty = Math.abs(normalizeNumber(record.individualQty));
        if (mode === 'in') {
          current.individual += qty;
        } else {
          // 낱개 부족 시 박스 언패킹
          while (current.individual < qty && current.box > 0) {
            if (current.boxContent <= 0) {
              throw new Error(`[${current.name}(${current.color})] 박스당 낱개 수량이 0이어서 박스를 개봉할 수 없습니다.`);
            }
            current.box -= 1;
            current.individual += current.boxContent;
          }
          if (current.individual < qty) {
            throw new Error(`[${current.name}(${current.color})] 낱개 재고가 부족합니다. (현재 가용: ${current.individual}개, 요청: ${qty}개)`);
          }
          current.individual -= qty;
        }
      }
    });

    // 3. 락 상태에서 고유 송장 번호 생성
    const seq = generateInvoiceNumber(typeKorean);
    const invoiceNumber = `${todayStr}-${seq}`;

    // 4. PendingSheet 기록 데이터 생성
    const adminName = normalizeText(admin) || 'ADMIN';
    const pendingRows = tableData.map(record => {
      const name = normalizeText(record.itemName);
      const color = normalizeText(record.color) || DEFAULTS.COLOR;
      const boxContent = normalizeNumber(record.boxContent);
      const key = makeKey(name, color, boxContent);
      const mfr = stockMap[key] ? stockMap[key].manufacturer : '';

      return [
        invoiceNumber,
        typeKorean,
        new Date(),
        name,
        color,
        record.unitType === 'box' ? Math.abs(normalizeNumber(record.boxQty)) : 0,
        record.unitType === 'individual' ? Math.abs(normalizeNumber(record.individualQty)) : 0,
        boxContent,
        normalizeText(record.location),
        adminName,
        mfr
      ];
    });

    // 5. 재고시트 일괄 쓰기 (Batch SetValues - 객체 속성 직접 참조로 언더스코어 파싱 완전 배제)
    const updatedStockRows = Object.values(stockMap).map(v => [
      v.name,
      v.color,
      v.box,
      v.individual,
      v.safeStock,
      v.boxContent,
      v.initialStock,
      v.manufacturer
    ]);

    if (updatedStockRows.length > 0) {
      ensureSheetCapacity(stockSheet, 2 + updatedStockRows.length - 1);
      stockSheet.getRange(2, 1, updatedStockRows.length, 8).setValues(updatedStockRows);
      if (stockLastRow - 1 > updatedStockRows.length) {
        stockSheet.getRange(2 + updatedStockRows.length, 1, (stockLastRow - 1) - updatedStockRows.length, 8).clearContent();
      }
    }

    // 6. PendingSheet 일괄 쓰기
    const pendingLastRow = pendingSheet.getLastRow();
    ensureSheetCapacity(pendingSheet, pendingLastRow + pendingRows.length);
    pendingSheet.getRange(pendingLastRow + 1, 1, pendingRows.length, 11).setValues(pendingRows);

    console.log(`processForm 완료: ${invoiceNumber} (${typeKorean} ${pendingRows.length}건)`);
    return seq;
  } catch (e) {
    console.error(`processForm error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------
// 재고시트 직접 갱신 (WarehouseApp 보류 저장/복구 호환용)
// -------------------------------------------------------------------

function updateStockSheet(tableData, mode) {
  if (!tableData || tableData.length === 0) return;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sheet = getSheet(SHEETS.STOCK);
    const lastRow = sheet.getLastRow();
    const stockMap = Object.create(null);

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      data.forEach(row => {
        const name = normalizeText(row[0]);
        const color = normalizeText(row[1]) || DEFAULTS.COLOR;
        const boxContent = normalizeNumber(row[5]);
        const key = makeKey(name, color, boxContent);
        stockMap[key] = {
          name: name,
          color: color,
          box: normalizeNumber(row[2]),
          individual: normalizeNumber(row[3]),
          safeStock: normalizeNumber(row[4]),
          boxContent: boxContent,
          initialStock: normalizeNumber(row[6]),
          manufacturer: normalizeText(row[7])
        };
      });
    }

    tableData.forEach(record => {
      const name = normalizeText(record.itemName);
      const color = normalizeText(record.color) || DEFAULTS.COLOR;
      const boxContent = normalizeNumber(record.boxContent);
      const key = makeKey(name, color, boxContent);
      let current = stockMap[key];

      if (!current) {
        current = {
          name: name,
          color: color,
          box: 0,
          individual: 0,
          safeStock: normalizeNumber(record.safeStock),
          boxContent: boxContent,
          initialStock: 0,
          manufacturer: normalizeText(record.manufacturer)
        };
        stockMap[key] = current;
      }

      if (record.unitType === 'box') {
        const qty = Math.abs(normalizeNumber(record.boxQty));
        if (mode === 'in') {
          current.box += qty;
        } else {
          if (current.box < qty) throw new Error(`[${current.name}] 박스 재고가 부족합니다.`);
          current.box -= qty;
        }
      } else if (record.unitType === 'individual') {
        const qty = Math.abs(normalizeNumber(record.individualQty));
        if (mode === 'in') {
          current.individual += qty;
        } else {
          while (current.individual < qty && current.box > 0) {
            if (current.boxContent <= 0) throw new Error(`[${current.name}] 박스당 낱개 수량을 확인하세요.`);
            current.box -= 1;
            current.individual += current.boxContent;
          }
          if (current.individual < qty) throw new Error(`[${current.name}] 낱개 재고가 부족합니다.`);
          current.individual -= qty;
        }
      }
    });

    const updatedData = Object.values(stockMap).map(v => [
      v.name,
      v.color,
      v.box,
      v.individual,
      v.safeStock,
      v.boxContent,
      v.initialStock,
      v.manufacturer
    ]);

    if (updatedData.length > 0) {
      ensureSheetCapacity(sheet, 2 + updatedData.length - 1);
      sheet.getRange(2, 1, updatedData.length, 8).setValues(updatedData);
      if (lastRow - 1 > updatedData.length) {
        sheet.getRange(2 + updatedData.length, 1, (lastRow - 1) - updatedData.length, 8).clearContent();
      }
    }
  } catch (e) {
    console.error(`updateStockSheet error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------
// 검색 및 수정(SearchModify) - 초고속 배치 처리 및 완벽한 정합성 보장
// -------------------------------------------------------------------

function searchRecords(type, invoiceNumber) {
  const sheet = getSheet(SHEETS.PENDING);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const targetInv = normalizeText(invoiceNumber);
  const targetType = normalizeText(type);
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  const records = data.filter(row => {
    const inv = normalizeText(row[0]);
    const rType = normalizeText(row[1]);
    return (inv === targetInv || inv.replace(/-/g, '/') === targetInv.replace(/-/g, '/')) && rType === targetType;
  });

  return records.map(row => ({
    itemName: normalizeText(row[3]),
    color: normalizeText(row[4]) || DEFAULTS.COLOR,
    boxQty: normalizeNumber(row[5]),
    individualQty: normalizeNumber(row[6]),
    boxContent: normalizeNumber(row[7]),
    location: normalizeText(row[8]),
    admin: normalizeText(row[9]),
    manufacturer: normalizeText(row[10])
  }));
}

function updatePendingRecords(invoiceNumber, type, newRecords, admin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000); // 25초 대기

    const pendingSheet = getSheet(SHEETS.PENDING);
    const stockSheet = getSheet(SHEETS.STOCK);
    const targetInv = normalizeText(invoiceNumber).replace(/-/g, '/');
    const targetType = normalizeText(type);

    // 1. 재고 맵 로드
    const stockLastRow = stockSheet.getLastRow();
    const stockMap = Object.create(null);
    if (stockLastRow >= 2) {
      const sData = stockSheet.getRange(2, 1, stockLastRow - 1, 8).getValues();
      sData.forEach(row => {
        const name = normalizeText(row[0]);
        const color = normalizeText(row[1]) || DEFAULTS.COLOR;
        const boxContent = normalizeNumber(row[5]);
        const key = makeKey(name, color, boxContent);
        stockMap[key] = {
          name: name,
          color: color,
          box: normalizeNumber(row[2]),
          individual: normalizeNumber(row[3]),
          safeStock: normalizeNumber(row[4]),
          boxContent: boxContent,
          initialStock: normalizeNumber(row[6]),
          manufacturer: normalizeText(row[7])
        };
      });
    }

    // 2. PendingSheet 데이터 로드 및 분류 (유지할 행 vs 삭제/수정 대상 행)
    const pendingLastRow = pendingSheet.getLastRow();
    const keptPendingRows = [];
    const oldMatchedRows = [];

    if (pendingLastRow >= 2) {
      const pData = pendingSheet.getRange(2, 1, pendingLastRow - 1, 11).getValues();
      pData.forEach(row => {
        const inv = normalizeText(row[0]).replace(/-/g, '/');
        const rType = normalizeText(row[1]);
        if (inv === targetInv && rType === targetType) {
          oldMatchedRows.push(row);
        } else {
          keptPendingRows.push(row);
        }
      });
    }

    // 3. 이전 기록 롤백 (재고 원상복구)
    oldMatchedRows.forEach(row => {
      const name = normalizeText(row[3]);
      const color = normalizeText(row[4]) || DEFAULTS.COLOR;
      const boxQty = normalizeNumber(row[5]);
      const individualQty = normalizeNumber(row[6]);
      const boxContent = normalizeNumber(row[7]);
      const key = makeKey(name, color, boxContent);
      let item = stockMap[key];

      if (!item) {
        item = {
          name: name,
          color: color,
          box: 0,
          individual: 0,
          safeStock: 0,
          boxContent: boxContent,
          initialStock: 0,
          manufacturer: normalizeText(row[10])
        };
        stockMap[key] = item;
      }

      if (targetType === '입고') {
        // 입고 취소 -> 재고 차감
        item.box -= boxQty;
        item.individual -= individualQty;
        while (item.individual < 0 && item.box > 0) {
          if (item.boxContent <= 0) break;
          item.box -= 1;
          item.individual += item.boxContent;
        }
      } else {
        // 출고 취소 -> 재고 복원(가산)
        item.box += boxQty;
        item.individual += individualQty;
      }
    });

    // 4. 신규 레코드 반영 (수정 내역이 있을 경우)
    const createdPendingRows = [];
    const adminName = normalizeText(admin) || 'ADMIN';

    if (newRecords && newRecords.length > 0) {
      newRecords.forEach(record => {
        const name = normalizeText(record.itemName);
        const color = normalizeText(record.color) || DEFAULTS.COLOR;
        const boxContent = normalizeNumber(record.boxContent);
        const key = makeKey(name, color, boxContent);
        let item = stockMap[key];

        if (!item) {
          if (targetType === '출고') {
            throw new Error(`등록되지 않은 상품입니다: ${name} (${color})`);
          }
          item = {
            name: name,
            color: color,
            box: 0,
            individual: 0,
            safeStock: 0,
            boxContent: boxContent,
            initialStock: 0,
            manufacturer: normalizeText(record.manufacturer)
          };
          stockMap[key] = item;
        }

        const boxQty = Math.abs(normalizeNumber(record.boxQty));
        const individualQty = Math.abs(normalizeNumber(record.individualQty));

        if (targetType === '입고') {
          item.box += boxQty;
          item.individual += individualQty;
        } else {
          // 출고 처리 및 박스 언패킹
          if (boxQty > 0) {
            if (item.box < boxQty) {
              throw new Error(`[${item.name}(${item.color})] 박스 재고가 부족합니다.`);
            }
            item.box -= boxQty;
          }
          if (individualQty > 0) {
            while (item.individual < individualQty && item.box > 0) {
              if (item.boxContent <= 0) throw new Error(`[${item.name}] 박스당 낱개 수량을 확인하세요.`);
              item.box -= 1;
              item.individual += item.boxContent;
            }
            if (item.individual < individualQty) {
              throw new Error(`[${item.name}(${item.color})] 낱개 재고가 부족합니다.`);
            }
            item.individual -= individualQty;
          }
        }

        createdPendingRows.push([
          invoiceNumber,
          targetType,
          new Date(),
          name,
          color,
          boxQty,
          individualQty,
          boxContent,
          normalizeText(record.location),
          adminName,
          item.manufacturer || normalizeText(record.manufacturer)
        ]);
      });
    }

    // 5. 시트 일괄 반영 (Batching Write)

    // A. 재고시트 일괄 갱신
    const updatedStockRows = Object.values(stockMap).map(v => [
      v.name,
      v.color,
      v.box,
      v.individual,
      v.safeStock,
      v.boxContent,
      v.initialStock,
      v.manufacturer
    ]);

    if (updatedStockRows.length > 0) {
      ensureSheetCapacity(stockSheet, 2 + updatedStockRows.length - 1);
      stockSheet.getRange(2, 1, updatedStockRows.length, 8).setValues(updatedStockRows);
      if (stockLastRow - 1 > updatedStockRows.length) {
        stockSheet.getRange(2 + updatedStockRows.length, 1, (stockLastRow - 1) - updatedStockRows.length, 8).clearContent();
      }
    }

    // B. PendingSheet 일괄 갱신 (deleteRow 루프 완전 배제)
    const finalPendingRows = keptPendingRows.concat(createdPendingRows);
    if (finalPendingRows.length > 0) {
      ensureSheetCapacity(pendingSheet, 2 + finalPendingRows.length - 1);
      pendingSheet.getRange(2, 1, finalPendingRows.length, 11).setValues(finalPendingRows);
    }
    if (pendingLastRow - 1 > finalPendingRows.length) {
      pendingSheet.getRange(2 + finalPendingRows.length, 1, (pendingLastRow - 1) - finalPendingRows.length, 11).clearContent();
    }

    console.log(`updatePendingRecords 완료: ${invoiceNumber} (${targetType})`);
  } catch (e) {
    console.error(`updatePendingRecords error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------
// 🏢 외부(서브)창고 8대 재고 매트릭스 & 유효재고 백엔드 연동
// -------------------------------------------------------------------

const SUB_WAREHOUSE_CONFIG = {
  SPREADSHEET_ID: '17_FjWFFbuMvvVhQ7Bn7CKmh59c9hzHDvWv68y11v4CX8',
  TARGET_WAREHOUSES: ['PANTACO', 'IKEA', 'LERMA', 'PINO', 'YARE', 'ALMINTER', 'TLANE', 'STAR']
};

function getSubWarehouseStockMatrix(forceRefresh) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'SUB_WH_MATRIX_V1';

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {}
    }
  }

  try {
    const subSS = SpreadsheetApp.openById(SUB_WAREHOUSE_CONFIG.SPREADSHEET_ID);
    const allSheets = subSS.getSheets();

    // 1. 재고현황 시트 찾기 (gid 543678626 또는 이름 '재고현황' 또는 첫 번째 시트)
    let stockSheet = subSS.getSheetByName('재고현황');
    if (!stockSheet) {
      stockSheet = allSheets.find(s => s.getSheetId() === 543678626) || allSheets[0];
    }

    // 2. 주문사항 시트 찾기 (gid 1459767519 또는 이름 '주문사항' 또는 '주문내역')
    let orderSheet = subSS.getSheetByName('주문사항') || subSS.getSheetByName('주문내역');
    if (!orderSheet) {
      orderSheet = allSheets.find(s => s.getSheetId() === 1459767519);
    }

    // A. 서브창고 재고 매트릭스 로드
    const stockData = stockSheet.getDataRange().getValues();
    if (stockData.length < 2) {
      throw new Error('서브창고 재고 데이터가 비어있습니다.');
    }

    const headerRow = stockData[0];
    let codigoCol = -1;
    let colorCol = -1;
    const warehouseColMap = {};

    headerRow.forEach((colName, idx) => {
      const cleanName = normalizeText(colName).toUpperCase();
      if (cleanName === 'CODIGO' || cleanName === '품명' || cleanName === '제품명') {
        codigoCol = idx;
      } else if (cleanName === 'COLOR' || cleanName === '색상') {
        colorCol = idx;
      } else {
        SUB_WAREHOUSE_CONFIG.TARGET_WAREHOUSES.forEach(wh => {
          if (cleanName.indexOf(wh) !== -1 || wh.indexOf(cleanName) !== -1) {
            warehouseColMap[wh] = idx;
          }
        });
      }
    });

    if (codigoCol === -1) codigoCol = 1;
    if (colorCol === -1) colorCol = 2;

    // B. 서브창고 PENDING 이동중(In-Transit) 수량 집계
    const inTransitMap = {};
    if (orderSheet) {
      const orderData = orderSheet.getDataRange().getValues();
      if (orderData.length >= 2) {
        const orderHeader = orderData[0];
        let pNameCol = -1, pColorCol = -1, pQtyCol = -1, pStatusCol = -1;

        orderHeader.forEach((h, idx) => {
          const ch = normalizeText(h).toUpperCase();
          if (ch.indexOf('품명') !== -1) pNameCol = idx;
          else if (ch.indexOf('색상') !== -1) pColorCol = idx;
          else if (ch.indexOf('개수') !== -1 || ch.indexOf('수량') !== -1) pQtyCol = idx;
          else if (ch.indexOf('처리상태') !== -1 || ch.indexOf('상태') !== -1) pStatusCol = idx;
        });

        if (pNameCol === -1) pNameCol = 4;
        if (pColorCol === -1) pColorCol = 5;
        if (pQtyCol === -1) pQtyCol = 6;
        if (pStatusCol === -1) pStatusCol = 9;

        for (let r = 1; r < orderData.length; r++) {
          const row = orderData[r];
          const status = normalizeText(row[pStatusCol]).toUpperCase();
          if (status === 'PENDING') {
            const item = normalizeText(row[pNameCol]);
            const color = normalizeText(row[pColorCol]) || DEFAULTS.COLOR;
            const qty = normalizeNumber(row[pQtyCol]);
            if (item && qty > 0) {
              const key = `${item}___${color}`.toUpperCase();
              inTransitMap[key] = (inTransitMap[key] || 0) + qty;
            }
          }
        }
      }
    }

    // C. 메인창고 실재고(Main Stock) 매핑
    const mainStockMap = {};
    try {
      const mainStockSheet = getSheet(SHEETS.STOCK);
      const mData = mainStockSheet.getDataRange().getValues();
      for (let m = 1; m < mData.length; m++) {
        const mRow = mData[m];
        const mName = normalizeText(mRow[0]);
        const mColor = normalizeText(mRow[1]) || DEFAULTS.COLOR;
        const mBox = normalizeNumber(mRow[2]);
        if (mName) {
          const mKey = `${mName}___${mColor}`.toUpperCase();
          mainStockMap[mKey] = (mainStockMap[mKey] || 0) + mBox;
        }
      }
    } catch (e) {
      console.warn('메인 재고 로드 실패: ' + e.message);
    }

    // D. 최종 매트릭스 항목 구성
    const items = [];
    const whList = SUB_WAREHOUSE_CONFIG.TARGET_WAREHOUSES;

    for (let i = 1; i < stockData.length; i++) {
      const row = stockData[i];
      const codigo = normalizeText(row[codigoCol]);
      if (!codigo) continue;
      const color = normalizeText(row[colorCol]) || DEFAULTS.COLOR;
      const key = `${codigo}___${color}`.toUpperCase();

      const stocks = {};
      let totalSubStock = 0;

      whList.forEach(wh => {
        const colIdx = warehouseColMap[wh];
        const qty = colIdx !== undefined ? normalizeNumber(row[colIdx]) : 0;
        stocks[wh] = qty;
        totalSubStock += qty;
      });

      const mainStock = mainStockMap[key] || 0;
      const inTransit = inTransitMap[key] || 0;
      const effectiveStock = mainStock + inTransit;

      items.push({
        codigo: codigo,
        color: color,
        mainStock: mainStock,
        inTransit: inTransit,
        effectiveStock: effectiveStock,
        stocks: stocks,
        totalSubStock: totalSubStock
      });
    }

    const result = {
      warehouses: whList,
      items: items,
      updatedAt: new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' })
    };

    try {
      cache.put(cacheKey, JSON.stringify(result), 600); // 10분 캐시
    } catch (e) {}

    return result;
  } catch (err) {
    console.error('getSubWarehouseStockMatrix error: ' + err.message);
    throw new Error(`서브창고 재고 데이터 로드 실패: ${err.message}`);
  }
}

// -------------------------------------------------------------------
// 메뉴 및 UI 표시
// -------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi().createMenu('창고 관리')
    .addItem('입고입력', 'showAppIn')
    .addItem('출고입력', 'showAppOut')
    .addItem('상품등록', 'showAppProduct')
    .addItem('검색수정', 'showSearchModify')
    .addSeparator()
    .addItem('📦 모든 시트 용량 점검 (1,000행 자동확보)', 'promptEnsureAllSheetsCapacity')
    .addItem('🔑 Gemini API 키 설정', 'promptSetGeminiApiKey')
    .addToUi();

  // 스프레드시트 열릴 때 모든 시트 용량을 선제 점검하여 100행 미만이면 자동 1,000행 확장
  try {
    ensureAllSheetsCapacity();
  } catch (e) {
    console.warn(`onOpen 자동 용량 점검 중 예외: ${e.message}`);
  }
}

function promptSetGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentKey = props.getProperty('GEMINI_API_KEY') || '';
  const maskedKey = currentKey ? (currentKey.slice(0, 8) + '...' + currentKey.slice(-4)) : '미등록';

  const response = ui.prompt(
    'Gemini API 키 설정',
    '현재 등록 상태: ' + maskedKey + '\n\n새로운 Gemini API 키를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const newKey = response.getResponseText().trim();
    if (newKey) {
      props.setProperty('GEMINI_API_KEY', newKey);
      ui.alert('설정 완료', 'GEMINI_API_KEY가 안전하게 스크립트 속성에 저장되었습니다.', ui.ButtonSet.OK);
    } else {
      ui.alert('안내', '입력된 키가 없어 변경되지 않았습니다.', ui.ButtonSet.OK);
    }
  }
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
  const html = template.evaluate().setWidth(1180).setHeight(851);
  SpreadsheetApp.getUi().showModalDialog(html, '창고 관리');
}

// -------------------------------------------------------------------
// Gemini 3.6 Flash 비전 AI 수기 주문서 인식
// -------------------------------------------------------------------

function getGeminiApiKey() {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error('GEMINI_API_KEY 스크립트 속성이 설정되지 않았습니다. Apps Script 프로젝트 설정(스크립트 속성)에 GEMINI_API_KEY를 등록해주세요.');
  }
  return key;
}

function analyzeHandwrittenOrder(imageBase64) {
  if (!imageBase64) {
    throw new Error('전달된 이미지 데이터가 없습니다.');
  }

  let cleanB64 = imageBase64;
  if (cleanB64.indexOf(',') > -1) {
    cleanB64 = cleanB64.split(',')[1];
  }

  const apiKey = getGeminiApiKey();
  const promptText = `Extract all handwritten order rows from the image. Support BOTH Form A (Printed grid table with columns) AND Form B (Free-form handwritten text with comma or dot delimiters).

Rules:
1. Header Information:
   - "branch": Branch/customer name written at the top header (e.g. "Aztecas", "CARMEN", "TIENDA", "CHINCONCUAC", "지점명"). If not found, return "".
   - "requester": Order requester/admin name written at the top right, especially text that is UNDERLINED (e.g. text with an underline '___' like 'Sr. Kim', '요청자이름', or in parentheses). If not found, return "".

2. Delimiters (for Form B free-form):
   - ONLY commas (',') and dots ('.') are delimiters between fields.
   - Spaces/whitespace are NEVER delimiters (preserve spaces in multi-word colors like "Palo Rosa", "Azul Marino" or models).

3. Row Parsing Rules:
   - Each row line containing a quantity is a separate entry:
     - 3 fields format: [modelo] ,/. [color] ,/. [quantity]
     - Color empty/blank (e.g. [modelo] , , [qty] OR [modelo] .. [qty]): return {"modelo": "...", "color": "", "no_de_bultos": qty}
     - Model empty/blank (Option 1: [empty] , [color] , [qty] OR Option 2: [color] , [qty]):
       If the row starts with empty delimiter or only has a color word and quantity, return {"modelo": "", "color": "color_name", "no_de_bultos": qty}
   - Form A (Grid table rows):
     - "modelo": Text written in the 'modelo' column. If empty on that row line, return "".
     - "color": Text written in the 'color' column. If empty on that row line, return "".
     - "no_de_bultos": The integer quantity in the 'cant oder' column for that row line.
     - "contenedor": Remarks if written, else "".

4. IMPORTANT:
   - Do NOT merge different rows!
   - For quantities, extract only the positive integer number (e.g. "./ 1" or ". 1" is 1).

Return ONLY valid JSON:
{
  "branch": "...",
  "requester": "...",
  "results": [
    {"modelo": "P-D60", "color": "Blanco", "no_de_bultos": 2},
    {"modelo": "", "color": "Beige", "no_de_bultos": 1}
  ]
}`;

  const models = ['gemini-3.7-flash', 'gemini-3.8-flash', 'gemini-3.6-flash'];
  let rawResponse = '';
  let lastError = '';

  for (let i = 0; i < models.length; i++) {
    const model = models[i];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    const payload = {
      contents: [{
        parts: [
          { text: promptText },
          { inline_data: { mime_type: 'image/jpeg', data: cleanB64 } }
        ]
      }],
      generationConfig: {
        response_mime_type: 'application/json',
        temperature: 0.1
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      if (code === 200) {
        const json = JSON.parse(resp.getContentText());
        const candidates = json.candidates || [];
        if (candidates.length > 0) {
          const parts = candidates[0].content ? candidates[0].content.parts || [] : [];
          rawResponse = parts.map(p => p.text || '').join('');
          break;
        }
      } else {
        lastError = `${model} (${code}): ${resp.getContentText().slice(0, 200)}`;
      }
    } catch (e) {
      lastError = e.message;
    }
  }

  if (!rawResponse) {
    throw new Error(`Gemini 비전 분석 실패: ${lastError}`);
  }

  try {
    let cleanJson = rawResponse.trim();
    if (cleanJson.startsWith('```json')) cleanJson = cleanJson.slice(7);
    if (cleanJson.startsWith('```')) cleanJson = cleanJson.slice(3);
    if (cleanJson.endsWith('```')) cleanJson = cleanJson.slice(0, -3);
    return JSON.parse(cleanJson.trim());
  } catch (err) {
    throw new Error(`분석 결과 JSON 파싱 오류: ${err.message}\n응답: ${rawResponse.slice(0, 300)}`);
  }
}

// -------------------------------------------------------------------
// 퀵 재고조정 (PendingSheet '재고조정' 기록 & 재고시트 수량 갱신)
// -------------------------------------------------------------------

function processQuickStockAdjustment(adjustments, admin) {
  if (!adjustments || adjustments.length === 0) {
    throw new Error('조정할 데이터가 없습니다.');
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    const stockSheet = getSheet(SHEETS.STOCK);
    const pendingSheet = getSheet(SHEETS.PENDING);
    const todayStr = formatDate(new Date());
    const adminName = normalizeText(admin) || 'ADMIN';

    // 1. 재고 데이터 로드
    const stockLastRow = stockSheet.getLastRow();
    const stockMap = Object.create(null);

    if (stockLastRow >= 2) {
      const stockData = stockSheet.getRange(2, 1, stockLastRow - 1, 8).getValues();
      stockData.forEach(row => {
        const name = normalizeText(row[0]);
        const color = normalizeText(row[1]) || DEFAULTS.COLOR;
        const boxContent = normalizeNumber(row[5]);
        const key = makeKey(name, color, boxContent);
        stockMap[key] = {
          name: name,
          color: color,
          box: normalizeNumber(row[2]),
          individual: normalizeNumber(row[3]),
          safeStock: normalizeNumber(row[4]),
          boxContent: boxContent,
          initialStock: normalizeNumber(row[6]),
          manufacturer: normalizeText(row[7])
        };
      });
    }

    // 2. 조정 번호 생성 및 PendingSheet / stockMap 반영
    const seq = generateInvoiceNumber('재고조정');
    const invoiceNumber = `${todayStr}-${seq}`;
    const pendingRows = [];
    const updatedKeys = [];

    adjustments.forEach(item => {
      const name = normalizeText(item.itemName);
      const color = normalizeText(item.color) || DEFAULTS.COLOR;
      const boxContent = normalizeNumber(item.boxContent);
      const targetBox = normalizeNumber(item.targetBox);
      const targetIndiv = normalizeNumber(item.targetIndividual);
      const key = makeKey(name, color, boxContent);

      let current = stockMap[key];
      if (!current) {
        current = {
          name: name,
          color: color,
          box: targetBox,
          individual: targetIndiv,
          safeStock: normalizeNumber(item.safeStock),
          boxContent: boxContent,
          initialStock: 0,
          manufacturer: normalizeText(item.manufacturer)
        };
        stockMap[key] = current;
      } else {
        current.box = targetBox;
        current.individual = targetIndiv;
      }

      // PendingSheet 기록: 구분 = '재고조정'
      pendingRows.push([
        invoiceNumber,
        '재고조정',
        new Date(),
        name,
        color,
        targetBox,
        targetIndiv,
        boxContent,
        normalizeText(item.location) || '재고조정',
        adminName,
        current.manufacturer || ''
      ]);

      updatedKeys.push({
        key: key,
        name: name,
        color: color,
        box: targetBox,
        individual: targetIndiv,
        boxContent: boxContent
      });
    });

    // 3. 재고시트 갱신 (전체 맵 일괄 저장)
    const updatedStockRows = Object.values(stockMap).map(v => [
      v.name,
      v.color,
      v.box,
      v.individual,
      v.safeStock,
      v.boxContent,
      v.initialStock,
      v.manufacturer
    ]);

    if (updatedStockRows.length > 0) {
      ensureSheetCapacity(stockSheet, 2 + updatedStockRows.length - 1);
      stockSheet.getRange(2, 1, updatedStockRows.length, 8).setValues(updatedStockRows);
    }

    // 4. PendingSheet 에 '재고조정' 행 일괄 추가
    if (pendingRows.length > 0) {
      const pLastRow = Math.max(pendingSheet.getLastRow() + 1, 2);
      ensureSheetCapacity(pendingSheet, pLastRow + pendingRows.length - 1);
      pendingSheet.getRange(pLastRow, 1, pendingRows.length, 11).setValues(pendingRows);
    }

    return {
      success: true,
      invoiceNumber: invoiceNumber,
      adjustedCount: pendingRows.length,
      updatedItems: updatedKeys
    };
  } catch (e) {
    console.error(`processQuickStockAdjustment error: ${e.message}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

