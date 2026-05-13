// ============================================================
// 무무즈에센셜 생산관리 시스템 — 초기 세팅 스크립트
// 실행 방법: setup() 함수 선택 후 실행 버튼 클릭 (1회만)
// ============================================================

const SHARED_DRIVE_ID = PropertiesService.getScriptProperties().getProperty('SHARED_DRIVE_ID');

// ── 아이템 유형 정의 ──────────────────────────────────────
const ITEM_TYPES = {
  TOP:    '상의',
  BOTTOM: '하의',
  OUTER:  '아우터',
  SET:    '세트업',
  DRESS:  '원피스'
};

// ── 각 유형별 사이즈 스펙 항목 (수정 가능) ────────────────
const SPEC_FIELDS = {
  TOP:    ['총장', '어깨너비', '가슴둘레', '소매길이', '밑단둘레'],
  BOTTOM: ['총장', '허리둘레', '엉덩이둘레', '밑위', '밑단둘레', '허벅지둘레'],
  OUTER:  ['총장', '어깨너비', '가슴둘레', '소매길이', '밑단둘레'],
  SET:    ['상의총장', '상의가슴', '상의소매', '하의총장', '하의허리', '하의밑위'],
  DRESS:  ['총장', '어깨너비', '가슴둘레', '허리둘레', '밑단둘레']
};

// ============================================================
// 메인 세팅 함수 — 이것만 실행하면 됩니다
// ============================================================
function setup() {
  console.log('=== 무무즈에센셜 생산관리 시스템 세팅 시작 ===');

  // 1. 폴더 구조 생성
  console.log('[1/5] 폴더 구조 생성 중...');
  const folders = createFolderStructure();

  // 2. 마스터 DB Sheets 생성
  console.log('[2/5] 마스터 DB 생성 중...');
  const masterSSId = createMasterDB(folders.masterDir);

  // 3. 운영 파일 Sheets 생성
  console.log('[3/5] 운영 파일 생성 중...');
  const opsSSId = createOperationFile(folders.opsDir);

  // 4. 작업지시서 Docs 템플릿 생성
  console.log('[4/5] 작업지시서 템플릿 생성 중...');
  const templateIds = createWorkOrderTemplates(folders.tmplDir);

  // 5. 시스템 설정값 기록
  console.log('[5/5] 시스템 설정 기록 중...');
  saveSystemConfig(masterSSId, opsSSId, folders, templateIds);

  console.log('=== 세팅 완료! ===');
  console.log('마스터 DB ID: ' + masterSSId);
  console.log('운영 파일 ID: ' + opsSSId);

  SpreadsheetApp.getUi().alert(
    '세팅 완료!\n\n' +
    '공유 드라이브에 다음 파일이 생성되었습니다:\n' +
    '- 01_마스터DB / 마스터DB.xlsx\n' +
    '- 02_운영파일 / 운영파일.xlsx\n' +
    '- 03_작업지시서템플릿 / WO_템플릿_* (5개)\n' +
    '- 04_작업지시서출력 (빈 폴더)\n' +
    '- 05_스와치이미지 (빈 폴더)\n\n' +
    '마스터DB > 시스템설정 탭에서 모든 파일 ID를 확인할 수 있습니다.'
  );
}


// ============================================================
// 1. 폴더 구조 생성
// ============================================================
function createFolderStructure() {
  const root = DriveApp.getFolderById(SHARED_DRIVE_ID);

  const masterDir  = getOrCreateFolder(root, '01_마스터DB');
  const opsDir     = getOrCreateFolder(root, '02_운영파일');
  const tmplDir    = getOrCreateFolder(root, '03_작업지시서템플릿');
  const outputDir  = getOrCreateFolder(root, '04_작업지시서출력');
  const swatchDir  = getOrCreateFolder(root, '05_스와치이미지');

  console.log('폴더 구조 생성 완료');
  return { masterDir, opsDir, tmplDir, outputDir, swatchDir };
}

function getOrCreateFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) {
    console.log('기존 폴더 사용: ' + name);
    return existing.next();
  }
  console.log('폴더 생성: ' + name);
  return parent.createFolder(name);
}


// ============================================================
// 2. 마스터 DB 생성
// ============================================================
function createMasterDB(folder) {
  const ss = SpreadsheetApp.create('마스터DB');
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);

  // 기본 시트 제거
  const defaultSheet = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');

  // ── 상품마스터 ─────────────────────────────────────────
  const productSheet = ss.insertSheet('상품마스터');
  const productHeaders = [
    '상품코드', '상품명', '카테고리', '아이템유형',
    '생산국가', '담당공장', '시즌', '담당MD', '등록일', '비고'
  ];
  setHeader(productSheet, productHeaders, '#E8F0FE');
  productSheet.setFrozenRows(1);
  productSheet.setColumnWidth(1, 100);
  productSheet.setColumnWidth(2, 180);
  addDataValidation(productSheet, 'C2:C1000', ['BABY', 'KIDS']);
  addDataValidation(productSheet, 'D2:D1000', Object.keys(ITEM_TYPES));
  addDataValidation(productSheet, 'E2:E1000', ['국내', '중국', '베트남']);

  // ── BOM (원부자재 소요량 기준) ─────────────────────────
  const bomSheet = ss.insertSheet('BOM');
  const bomHeaders = [
    '상품코드', '자재코드', '자재명', '단위', '단위소요량',
    '로스율(%)', '실소요량', '공급업체코드', '비고'
  ];
  setHeader(bomSheet, bomHeaders, '#E6F4EA');
  bomSheet.setFrozenRows(1);
  // 실소요량 = 단위소요량 × (1 + 로스율/100) 수식 안내 메모
  bomSheet.getRange('G1').setNote('수식: =E2*(1+F2/100)');
  addDataValidation(bomSheet, 'D2:D1000', ['m', 'y', 'EA', 'SET', 'cm', 'g', 'kg']);

  // ── 자재마스터 ─────────────────────────────────────────
  const matSheet = ss.insertSheet('자재마스터');
  const matHeaders = [
    '자재코드', '자재명', '구분', '단위', '단가', '공급업체코드', '리드타임(일)', '비고'
  ];
  setHeader(matSheet, matHeaders, '#FFF3E0');
  matSheet.setFrozenRows(1);
  addDataValidation(matSheet, 'C2:C1000', ['원단', '부자재', '라벨', '포장재', '기타']);

  // ── 업체마스터 ─────────────────────────────────────────
  const vendorSheet = ss.insertSheet('업체마스터');
  const vendorHeaders = [
    '업체코드', '업체명', '구분', '담당자', '연락처',
    '이메일', '카카오ID', '국가', '비고'
  ];
  setHeader(vendorSheet, vendorHeaders, '#F3E5F5');
  vendorSheet.setFrozenRows(1);
  addDataValidation(vendorSheet, 'C2:C1000', ['공장', '원단', '부자재', '기타']);

  // ── 사이즈스펙 ─────────────────────────────────────────
  // 상품마다 사이즈가 다르므로 유연한 구조로 설계
  // 구조: 상품코드 | 스펙항목 | 사이즈1 | 값1 | 사이즈2 | 값2 | ...
  const sizeSheet = ss.insertSheet('사이즈스펙');
  const sizeHeaders = [
    '상품코드', '스펙항목',
    '사이즈1', '값1', '사이즈2', '값2', '사이즈3', '값3',
    '사이즈4', '값4', '사이즈5', '값5', '사이즈6', '값6',
    '사이즈7', '값7', '사이즈8', '값8'
  ];
  setHeader(sizeSheet, sizeHeaders, '#E0F7FA');
  sizeSheet.setFrozenRows(1);
  sizeSheet.setFrozenColumns(2);
  // 샘플 데이터 (구조 이해용)
  const sampleSize = [
    ['MMZ-SAMPLE', '총장', '50', '38', '60', '40', '70', '42', '80', '44', '', '', '', '', '', '', '', ''],
    ['MMZ-SAMPLE', '가슴둘레', '50', '48', '60', '50', '70', '52', '80', '54', '', '', '', '', '', '', '', ''],
  ];
  sizeSheet.getRange(2, 1, sampleSize.length, sampleSize[0].length).setValues(sampleSize);
  sizeSheet.getRange(2, 1, sampleSize.length, sizeHeaders.length)
    .setBackground('#F8F9FA').setFontColor('#999999');

  // ── 시스템설정 (나중에 채워짐) ─────────────────────────
  const configSheet = ss.insertSheet('시스템설정');
  setHeader(configSheet, ['설정키', '값', '설명'], '#ECEFF1');
  configSheet.setColumnWidth(1, 200);
  configSheet.setColumnWidth(2, 320);
  configSheet.setColumnWidth(3, 300);

  // 기본 시트 삭제
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  console.log('마스터 DB 생성 완료: ' + ss.getId());
  return ss.getId();
}


// ============================================================
// 3. 운영 파일 생성
// ============================================================
function createOperationFile(folder) {
  const ss = SpreadsheetApp.create('운영파일');
  DriveApp.getFileById(ss.getId()).moveTo(folder);

  const defaultSheet = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');

  // ── 작업지시서목록 ─────────────────────────────────────
  const woSheet = ss.insertSheet('작업지시서목록');
  const woHeaders = [
    'WO번호', '상품코드', '상품명', '아이템유형', '카테고리',
    '시즌', '발행일', '납기일', '상태', '생산국가', '담당공장',
    'PDF링크', '담당자', '비고'
  ];
  setHeader(woSheet, woHeaders, '#E8F0FE');
  woSheet.setFrozenRows(1);
  woSheet.setFrozenColumns(2);
  addDataValidation(woSheet, 'I2:I1000',
    ['초안', '발행완료', '생산중', '납품완료', '취소']);
  addDataValidation(woSheet, 'D2:D1000', Object.keys(ITEM_TYPES));
  addDataValidation(woSheet, 'E2:E1000', ['BABY', 'KIDS']);

  // ── 발주수량 ───────────────────────────────────────────
  // 사이즈가 상품마다 달라서 WO번호+상품코드 기준으로 사이즈별 행 구조
  const qtySheet = ss.insertSheet('발주수량');
  const qtyHeaders = [
    'WO번호', '상품코드', '사이즈', '발주수량', '리오더여부',
    '원발주WO', '입력일시', '입력자'
  ];
  setHeader(qtySheet, qtyHeaders, '#E6F4EA');
  qtySheet.setFrozenRows(1);
  addDataValidation(qtySheet, 'E2:E1000', ['신규', '리오더']);

  // ── 원부자재소요량 ─────────────────────────────────────
  const matReqSheet = ss.insertSheet('원부자재소요량');
  const matReqHeaders = [
    'WO번호', '상품코드', '자재코드', '자재명', '단위',
    '단위소요량', '총발주수량', '총소요량', '로스율(%)', '실소요량',
    '공급업체코드', '업체명', '승인여부', '승인일시'
  ];
  setHeader(matReqSheet, matReqHeaders, '#FFF3E0');
  matReqSheet.setFrozenRows(1);
  matReqSheet.setFrozenColumns(2);
  addDataValidation(matReqSheet, 'M2:M1000', ['대기', '승인', '반려']);
  // 총소요량·실소요량은 자동계산 안내 메모
  matReqSheet.getRange('H1').setNote('= 단위소요량 × 총발주수량');
  matReqSheet.getRange('J1').setNote('= 총소요량 × (1 + 로스율/100)');

  // ── 리오더이력 ─────────────────────────────────────────
  const reorderSheet = ss.insertSheet('리오더이력');
  const reorderHeaders = [
    '리오더WO', '원본WO', '상품코드', '상품명',
    '리오더일자', '총수량', '담당MD', '사유', '상태'
  ];
  setHeader(reorderSheet, reorderHeaders, '#F3E5F5');
  reorderSheet.setFrozenRows(1);

  if (defaultSheet) ss.deleteSheet(defaultSheet);

  console.log('운영 파일 생성 완료: ' + ss.getId());
  return ss.getId();
}


// ============================================================
// 4. 작업지시서 Docs 템플릿 생성
// ============================================================
function createWorkOrderTemplates(folder) {
  const templateIds = {};

  Object.keys(ITEM_TYPES).forEach(function(typeKey) {
    const typeName = ITEM_TYPES[typeKey];
    const docName  = 'WO_템플릿_' + typeKey + '_' + typeName;
    const doc      = DocumentApp.create(docName);
    DriveApp.getFileById(doc.getId()).moveTo(folder);

    buildWorkOrderTemplate(doc, typeKey, typeName);
    templateIds[typeKey] = doc.getId();
    console.log('템플릿 생성: ' + docName + ' (' + doc.getId() + ')');
  });

  return templateIds;
}

function buildWorkOrderTemplate(doc, typeKey, typeName) {
  const body = doc.getBody();
  body.clear();

  // ── 스타일 설정 ─────────────────────────────────────────
  const titleStyle = {};
  titleStyle[DocumentApp.Attribute.FONT_SIZE]  = 18;
  titleStyle[DocumentApp.Attribute.BOLD]        = true;
  titleStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;

  const h2Style = {};
  h2Style[DocumentApp.Attribute.FONT_SIZE]      = 11;
  h2Style[DocumentApp.Attribute.BOLD]           = true;

  // ── 헤더 영역 ───────────────────────────────────────────
  const title = body.appendParagraph('MUMUZ ESSENTIAL');
  title.setAttributes(titleStyle);

  const subtitle = body.appendParagraph('작업지시서 (' + typeName + ')   Work Order');
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subtitle.setFontSize(10);

  body.appendHorizontalRule();

  // WO 메타 정보 테이블
  const metaTable = body.appendTable([
    ['WO No.', '{{WO_NO}}',     '발행일',   '{{ISSUE_DATE}}'],
    ['시즌',    '{{SEASON}}',    '납기일',   '{{DUE_DATE}}'],
    ['담당MD',  '{{MD_NAME}}',   '담당디자이너', '{{DESIGNER_NAME}}'],
  ]);
  styleTable(metaTable, '#F8F9FA');

  // ── 기본 정보 ───────────────────────────────────────────
  appendSectionHeader(body, '기본 정보', h2Style);
  const basicTable = body.appendTable([
    ['상품코드',  '{{PRODUCT_CODE}}',  '상품명',   '{{PRODUCT_NAME}}'],
    ['카테고리',  '{{CATEGORY}}',      '아이템유형', typeName],
    ['생산국가',  '{{COUNTRY}}',       '담당공장',  '{{FACTORY}}'],
  ]);
  styleTable(basicTable, '#FFFFFF');

  // ── 원단 정보 + 스와치 ──────────────────────────────────
  appendSectionHeader(body, '원단 정보', h2Style);
  const fabricTable = body.appendTable([
    ['원단명',  '{{FABRIC_NAME}}',  '조성',    '{{FABRIC_COMP}}'],
    ['색상',    '{{COLOR}}',        '공급처',  '{{FABRIC_VENDOR}}'],
    ['스와치',  '{{SWATCH_IMG}}',   '',        ''],
  ]);
  styleTable(fabricTable, '#FFFFFF');
  // 스와치 셀 병합 (행2, 열1~4)
  fabricTable.getRow(2).getCell(1).setWidth(400);

  // ── 사이즈 스펙 (유형별) ────────────────────────────────
  appendSectionHeader(body, '사이즈 스펙', h2Style);
  const specFields = SPEC_FIELDS[typeKey];
  // 헤더행: 스펙항목 | 사이즈1 | 값1 | 사이즈2 | 값2 | ...
  const specHeaderRow = ['스펙항목', '사이즈', '값', '사이즈', '값', '사이즈', '값', '사이즈', '값'];
  const specRows      = [specHeaderRow];
  specFields.forEach(function(field) {
    specRows.push([
      field,
      '{{SZ_1}}', '{{VAL_' + field + '_1}}',
      '{{SZ_2}}', '{{VAL_' + field + '_2}}',
      '{{SZ_3}}', '{{VAL_' + field + '_3}}',
      '{{SZ_4}}', '{{VAL_' + field + '_4}}'
    ]);
  });
  const specTable = body.appendTable(specRows);
  styleTable(specTable, '#E0F7FA');
  specTable.getRow(0).editAsText().setBold(true);

  // ── 발주 수량 ───────────────────────────────────────────
  appendSectionHeader(body, '발주 수량', h2Style);
  const qtyHeaderRow = ['사이즈', '수량', '사이즈', '수량', '사이즈', '수량', '사이즈', '수량', '합계'];
  const qtyValueRow  = [
    '{{QTY_SZ1}}', '{{QTY_1}}',
    '{{QTY_SZ2}}', '{{QTY_2}}',
    '{{QTY_SZ3}}', '{{QTY_3}}',
    '{{QTY_SZ4}}', '{{QTY_4}}',
    '{{QTY_TOTAL}}'
  ];
  const qtyTable = body.appendTable([qtyHeaderRow, qtyValueRow]);
  styleTable(qtyTable, '#E6F4EA');
  qtyTable.getRow(0).editAsText().setBold(true);

  // ── 원부자재 소요량 (자동계산) ──────────────────────────
  appendSectionHeader(body, '원부자재 소요량  ※ 자동계산', h2Style);
  const bomHeaderRow = ['자재명', '단위', '단위소요량', '발주수량', '총소요량', '공급업체'];
  // 실제 생성 시 BOM 데이터 기반으로 행이 동적 추가됨
  const bomPlaceholder = ['{{BOM_ROWS}}', '', '', '', '', ''];
  const bomTable = body.appendTable([bomHeaderRow, bomPlaceholder]);
  styleTable(bomTable, '#FFF3E0');
  bomTable.getRow(0).editAsText().setBold(true);

  // ── 카테고리별 조건 항목 ─────────────────────────────────
  appendSectionHeader(body, 'KC 인증 (베이비 전용)', h2Style);
  const kcTable = body.appendTable([
    ['KC 인증번호', '{{KC_CERT}}', '안전기준', '{{SAFETY_STD}}'],
    ['봉제방식',    '{{STITCH}}',  '기타',     '{{KC_ETC}}'],
  ]);
  styleTable(kcTable, '#FFF9C4');
  const kcNote = body.appendParagraph('※ KIDS 아이템은 이 섹션을 삭제하세요');
  kcNote.setFontSize(9).setItalic(true);
  kcNote.editAsText().setForegroundColor('#999999');

  // ── 추가 지시사항 ────────────────────────────────────────
  appendSectionHeader(body, '디자인 지시 / 특이사항', h2Style);
  const noteTable = body.appendTable([
    ['{{DESIGN_NOTE}}']
  ]);
  noteTable.getRow(0).setMinimumRowHeight(80);
  styleTable(noteTable, '#FFFFFF');

  // ── 결재란 ───────────────────────────────────────────────
  body.appendHorizontalRule();
  appendSectionHeader(body, '결재', h2Style);
  const approvalTable = body.appendTable([
    ['디자이너', 'MD', '생산', '승인'],
    ['{{SIGN_DESIGNER}}', '{{SIGN_MD}}', '{{SIGN_PROD}}', '{{SIGN_APPR}}'],
    ['{{DATE_DESIGNER}}', '{{DATE_MD}}', '{{DATE_PROD}}', '{{DATE_APPR}}'],
  ]);
  styleTable(approvalTable, '#FFFFFF');
  const approvalHeaderRow = approvalTable.getRow(0);
  approvalHeaderRow.editAsText().setBold(true);
  for (let ci = 0; ci < approvalHeaderRow.getNumCells(); ci++) {
    approvalHeaderRow.getCell(ci).setBackgroundColor('#ECEFF1');
  }
  approvalTable.getRow(1).setMinimumRowHeight(50);

  doc.saveAndClose();
}

function appendSectionHeader(body, text, style) {
  body.appendParagraph('');
  const p = body.appendParagraph('  ' + text);
  p.setAttributes(style);
  return p;
}

function styleTable(table, headerBg) {
  table.setBorderColor('#CCCCCC');
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4);
      cell.setPaddingBottom(4);
      cell.setPaddingLeft(6);
      cell.setPaddingRight(6);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    }
  }
}


// ============================================================
// 5. 시스템 설정값 저장
// ============================================================
function saveSystemConfig(masterSSId, opsSSId, folders, templateIds) {
  const ss          = SpreadsheetApp.openById(masterSSId);
  const configSheet = ss.getSheetByName('시스템설정');

  const configs = [
    ['SHARED_DRIVE_ID',   SHARED_DRIVE_ID,              '공유 드라이브 ID'],
    ['MASTER_SS_ID',      masterSSId,                   '마스터DB Sheets ID'],
    ['OPS_SS_ID',         opsSSId,                      '운영파일 Sheets ID'],
    ['MASTER_DIR_ID',     folders.masterDir.getId(),    '01_마스터DB 폴더 ID'],
    ['OPS_DIR_ID',        folders.opsDir.getId(),       '02_운영파일 폴더 ID'],
    ['TMPL_DIR_ID',       folders.tmplDir.getId(),      '03_템플릿 폴더 ID'],
    ['OUTPUT_DIR_ID',     folders.outputDir.getId(),    '04_출력 폴더 ID'],
    ['SWATCH_DIR_ID',     folders.swatchDir.getId(),    '05_스와치 폴더 ID'],
    ['TMPL_ID_TOP',       templateIds['TOP'],            'Docs 템플릿: 상의'],
    ['TMPL_ID_BOTTOM',    templateIds['BOTTOM'],         'Docs 템플릿: 하의'],
    ['TMPL_ID_OUTER',     templateIds['OUTER'],          'Docs 템플릿: 아우터'],
    ['TMPL_ID_SET',       templateIds['SET'],            'Docs 템플릿: 세트업'],
    ['TMPL_ID_DRESS',     templateIds['DRESS'],          'Docs 템플릿: 원피스'],
    ['SETUP_DATE',        new Date().toLocaleString('ko-KR'), '세팅 실행 일시'],
  ];

  configSheet.getRange(2, 1, configs.length, 3).setValues(configs);

  // 설정 시트 보호 (실수로 수정 방지)
  const protection = configSheet.protect().setDescription('시스템설정 보호');
  protection.setWarningOnly(true);

  console.log('시스템 설정 저장 완료');
}


// ============================================================
// 헬퍼: 헤더 스타일 적용
// ============================================================
function setHeader(sheet, headers, bgColor) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground(bgColor || '#E8F0FE');
  range.setFontWeight('bold');
  range.setFontSize(11);
  range.setBorder(true, true, true, true, true, true,
    '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 32);

  // 열 너비 자동 조정
  for (let i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 120);
  }
}


// ============================================================
// 헬퍼: 드롭다운 데이터 유효성 검사
// ============================================================
function addDataValidation(sheet, rangeNotation, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(rangeNotation).setDataValidation(rule);
}


// ============================================================
// 설정값 읽기 헬퍼 — 다른 .gs 파일에서 공통 사용
// ============================================================
function getConfig(key) {
  const ss          = SpreadsheetApp.openById(getMasterSSId());
  const configSheet = ss.getSheetByName('시스템설정');
  const data        = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  throw new Error('설정값 없음: ' + key);
}

// setup() 실행 전에는 SHARED_DRIVE_ID로 마스터 SS를 찾아야 함
// setup() 실행 후에는 시스템설정에서 읽음
function getMasterSSId() {
  const files = DriveApp.getFolderById(SHARED_DRIVE_ID)
    .getFoldersByName('01_마스터DB').next()
    .getFilesByName('마스터DB');
  if (files.hasNext()) return files.next().getId();
  throw new Error('마스터DB 파일을 찾을 수 없습니다. setup()을 먼저 실행하세요.');
}
