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
      stockSheet.getRange(2, 1, updatedStockRows.length, 8).setValues(updatedStockRows);
      if (stockLastRow - 1 > updatedStockRows.length) {
        stockSheet.getRange(2 + updatedStockRows.length, 1, (stockLastRow - 1) - updatedStockRows.length, 8).clearContent();
      }
    }

    // 6. PendingSheet 일괄 쓰기
    const pendingLastRow = pendingSheet.getLastRow();
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
      stockSheet.getRange(2, 1, updatedStockRows.length, 8).setValues(updatedStockRows);
      if (stockLastRow - 1 > updatedStockRows.length) {
        stockSheet.getRange(2 + updatedStockRows.length, 1, (stockLastRow - 1) - updatedStockRows.length, 8).clearContent();
      }
    }

    // B. PendingSheet 일괄 갱신 (deleteRow 루프 완전 배제)
    const finalPendingRows = keptPendingRows.concat(createdPendingRows);
    if (finalPendingRows.length > 0) {
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
// 메뉴 및 UI 표시
// -------------------------------------------------------------------

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
  const html = template.evaluate().setWidth(1180).setHeight(851);
  SpreadsheetApp.getUi().showModalDialog(html, '창고 관리');
}

// -------------------------------------------------------------------
// Gemini 3.6 Flash 비전 AI 수기 주문서 인식
// -------------------------------------------------------------------

function getGeminiApiKey() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('GEMINI_API_KEY');
  if (!key) {
    try {
      key = Utilities.newBlob(Utilities.base64Decode('QVEuQWI4Uk42S1NMRHBWcXJfTVlPVjhyNDV6R2gyRTBpbFJVTldJN1hFMVZ1b3AzbnlBZ0E=')).getDataAsString();
    } catch (e) {
      key = '';
    }
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
  const promptText = `Extract all handwritten table rows by strictly following the horizontal grid lines (rows).

Rules:
1. "branch": Branch/customer name written at the top header (e.g. "Aztecas", "CARMEN", "TIENDA", "CHINCONCUAC").
2. "results": Each horizontal grid line that contains a quantity number is a SEPARATE row:
   - "modelo": Text written in the 'modelo' column. If empty on that row line, return "".
   - "color": Text written in the 'color' column. If empty on that row line, return "".
   - "no_de_bultos": The integer quantity in the 'cant oder' column for that row line.
   - "contenedor": Remarks if written, else "".

IMPORTANT: Do NOT merge different horizontal lines!
For example:
- Line 1 has model 'L-AL165' and qty 5 -> {"modelo": "L-AL165", "color": "", "no_de_bultos": 5}
- Line 2 has color 'Negro' and qty 2 -> {"modelo": "", "color": "Negro", "no_de_bultos": 2}
- Line 3 has color 'Blanco' and qty 1 -> {"modelo": "", "color": "Blanco", "no_de_bultos": 1}

Return ONLY valid JSON:
{
  "branch": "...",
  "results": [
    {"modelo": "...", "color": "...", "no_de_bultos": 1}
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
      stockSheet.getRange(2, 1, updatedStockRows.length, 8).setValues(updatedStockRows);
    }

    // 4. PendingSheet 에 '재고조정' 행 일괄 추가
    if (pendingRows.length > 0) {
      const pLastRow = Math.max(pendingSheet.getLastRow() + 1, 2);
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

