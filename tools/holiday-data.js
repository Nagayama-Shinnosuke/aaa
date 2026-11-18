const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');

const ROOT = path.join(__dirname, '..');
const EXCEL_DIR = path.join(ROOT, 'excel');
const HOLIDAYS = [
  {
    key: 'obon',
    label: 'お盆',
    file: 'obon.xlsx',
    columns: [
      { key: 'normalDelivery', header: 'お盆期間納期' },
      { key: 'deadline', header: '締切時間' },
      { key: 'period', header: 'お盆期間' },
      { key: 'beforeHolidayDeadline', header: '期間前締切' },
      { key: 'afterHolidayDeadline', header: '期間後締切' },
    ],
  },
  {
    key: 'goldenWeek',
    label: 'ゴールデンウィーク',
    file: 'goldenWeek.xlsx',
    columns: [
      { key: 'normalDelivery', header: 'ゴールデンウィーク期間納期' },
      { key: 'deadline', header: '締切時間' },
      { key: 'period', header: 'ゴールデンウィーク期間' },
      { key: 'beforeHolidayDeadline', header: '期間前締切' },
      { key: 'afterHolidayDeadline', header: '期間後締切' },
    ],
  },
  {
    key: 'yearEndNewYear',
    label: '年末年始',
    file: 'yearEndNewYear.xlsx',
    columns: [
      { key: 'normalDelivery', header: '通常納期' },
      { key: 'period', header: '年末年始期間' },
      { key: 'beforeHolidayDeadline', header: '年内に届く最終ご注文日' },
      { key: 'afterHolidayDeadline', header: '年初すぐに届くご注文日' },
    ],
  },
];

function getColumns(holiday) {
  return [
    { key: 'name', header: '商品名', width: 48 },
    ...holiday.columns.map(column => ({ ...column, width: 24 })),
    { key: 'onlyClub', header: 'clubのみ', width: 14 },
  ];
}

function requireString(value, label) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`${label} は空欄にできません。`);
  }
  return value;
}

function getCellValue(row, columnIndex) {
  const value = row.getCell(columnIndex).value;
  if (value && typeof value === 'object' && 'text' in value) {
    return value.text;
  }
  return value;
}

function readExcelString(row, columnIndex, label) {
  const value = getCellValue(row, columnIndex);
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`${label} は文字列で入力してください。`);
  }
  return value;
}

function readOnlyClub(value, label) {
  if (value == null || value === '') return false;
  if (value === '○') return true;
  throw new Error(`${label} は空欄または「○」で入力してください。`);
}

function loadSourceProducts() {
  const sourcePath = path.join(ROOT, 'products.json');
  const source = JSON.parse(fs.readFileSync(sourcePath, 'utf8'));
  if (!Array.isArray(source) || source.length === 0) {
    throw new Error('products.json に商品データがありません。');
  }
  return source;
}

async function loadExistingValues(holiday, source) {
  const inputPath = path.join(EXCEL_DIR, holiday.file);
  if (!fs.existsSync(inputPath)) return null;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);
  const sheet = workbook.getWorksheet('商品データ');
  if (!sheet) {
    throw new Error(`${holiday.file} に「商品データ」シートがありません。`);
  }
  const headerIndexes = new Map();
  sheet.getRow(1).eachCell((cell, columnNumber) => {
    headerIndexes.set(cell.value, columnNumber);
  });
  const oldHeaders = {
    name: '商品名',
    onlyClub: 'clubのみ',
    ...Object.fromEntries(holiday.columns.map(column => [column.key, column.header])),
  };
  const legacyHeaderAliases = holiday.key === 'yearEndNewYear'
    ? {
      beforeHolidayDeadline: '年内最終注文日',
      afterHolidayDeadline: '年初注文日',
      onlyClub: 'クラブページのみ',
    }
    : { onlyClub: 'クラブページのみ' };
  if (holiday.key === 'yearEndNewYear') {
    oldHeaders.deadline = '締切時間（JSON保持用）';
  }
  const nameColumn = headerIndexes.get(oldHeaders.name);
  const existingValues = [];

  for (let index = 0; index < source.length; index += 1) {
    const row = sheet.getRow(index + 2);
    if (!nameColumn || getCellValue(row, nameColumn) !== source[index].name) {
      throw new Error(
        `${holiday.file} の商品名・行順が products.json と異なるため、安全に更新できません。`,
      );
    }
    const values = {};
    for (const column of holiday.columns) {
      const oldColumn = headerIndexes.get(oldHeaders[column.key])
        || headerIndexes.get(legacyHeaderAliases[column.key]);
      if (!oldColumn) {
        throw new Error(`${holiday.file} に「${oldHeaders[column.key]}」列がありません。`);
      }
      values[column.key] = getCellValue(row, oldColumn);
    }
    const onlyClubColumn = headerIndexes.get(oldHeaders.onlyClub)
      || headerIndexes.get(legacyHeaderAliases.onlyClub);
    values.onlyClub = onlyClubColumn
      ? getCellValue(row, onlyClubColumn)
      : (source[index].onlyClub ? '○' : '');
    existingValues.push(values);
  }
  return existingValues;
}

async function createTemplates({ refresh = false } = {}) {
  const source = loadSourceProducts();
  fs.mkdirSync(EXCEL_DIR, { recursive: true });

  for (const holiday of HOLIDAYS) {
    const outputPath = path.join(EXCEL_DIR, holiday.file);
    if (fs.existsSync(outputPath) && !refresh) {
      throw new Error(`${outputPath} はすでに存在します。上書きしないため処理を中止しました。`);
    }
    const existingValues = refresh ? await loadExistingValues(holiday, source) : null;
    const columns = getColumns(holiday);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Holiday delivery data converter';
    workbook.subject = `${holiday.label}の商品納期データ`;
    const sheet = workbook.addWorksheet('商品データ');
    sheet.columns = columns.map(column => ({
      header: column.header,
      key: column.key,
      width: column.width,
      style: { numFmt: '@', alignment: { vertical: 'top', wrapText: true } },
    }));
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: source.length + 1, column: columns.length },
    };
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF24527A' },
    };

    source.forEach((product, index) => {
      const delivery = product[holiday.key];
      if (!delivery) {
        throw new Error(`${product.name} に ${holiday.key} のデータがありません。`);
      }
      const row = {
        name: product.name,
        onlyClub: product.onlyClub ? '○' : '',
        ...delivery,
        ...(existingValues ? existingValues[index] : {}),
      };
      sheet.addRow(columns.map(column => row[column.key] ?? ''));
      const onlyClubColumn = columns.findIndex(column => column.key === 'onlyClub') + 1;
      sheet.getCell(index + 2, onlyClubColumn).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"○"'],
      };
    });

    const guide = workbook.addWorksheet('入力方法');
    guide.columns = [{ width: 110 }];
    guide.addRows([
      [`${holiday.label}の納期データ入力用ファイルです。`],
      ['「商品データ」シートの2行目以降を編集してください。'],
      ['Excelには商品名と、サイトに表示する納期・期間・注文日だけを記載します。'],
      ['「clubのみ」は空欄ならfalse、○を入力するとtrueです。'],
      ['納期・期間・注文日は文字列で入力してください（例：12/25(金)、3営業日後）。'],
      ['商品名は変更せず、行の並び替えや削除をしないでください。商品情報との照合に使います。'],
      ['更新後、プロジェクトのルートで npm run holiday:convert を実行すると3種類のJSONを出力します。'],
    ]);
    guide.getColumn(1).alignment = { wrapText: true, vertical: 'top' };
    guide.getColumn(1).width = 110;
    guide.eachRow(row => { row.height = 30; });

    await workbook.xlsx.writeFile(outputPath);
    console.log(`${refresh ? '更新' : '作成'}: ${path.relative(ROOT, outputPath)} (${source.length} 商品)`);
  }
}

async function readHolidayWorkbook(holiday) {
  const inputPath = path.join(EXCEL_DIR, holiday.file);
  if (!fs.existsSync(inputPath)) {
    throw new Error(`${inputPath} がありません。`);
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);
  const sheet = workbook.getWorksheet('商品データ');
  if (!sheet) {
    throw new Error(`${holiday.file} に「商品データ」シートがありません。`);
  }

  const columns = getColumns(holiday);
  const expectedHeaders = columns.map(column => column.header);
  const actualHeaders = expectedHeaders.map((_, index) => sheet.getCell(1, index + 1).value);
  if (JSON.stringify(actualHeaders) !== JSON.stringify(expectedHeaders)) {
    throw new Error(`${holiday.file} の1行目の見出しがテンプレートと異なります。`);
  }

  const source = loadSourceProducts();
  const productMetadata = new Map();
  for (const product of source) {
    const metadata = { category: product.category };
    const existing = productMetadata.get(product.name);
    if (existing && existing.category !== metadata.category) {
      throw new Error(`${product.name} のカテゴリーが商品データ内で一致しません。`);
    }
    productMetadata.set(product.name, metadata);
  }

  const columnIndexes = new Map(columns.map((column, index) => [column.key, index + 1]));
  const products = [];
  for (let index = 0; index < source.length; index += 1) {
    const rowNumber = index + 2;
    const row = sheet.getRow(rowNumber);
    const label = `${holiday.file} ${rowNumber}行目`;
    const name = requireString(
      getCellValue(row, columnIndexes.get('name')),
      `${label}の商品名`,
    );
    if (name !== source[index].name) {
      throw new Error(`${label}の商品名・行順が products.json と一致しません。`);
    }
    const metadata = productMetadata.get(name);
    const onlyClub = readOnlyClub(
      getCellValue(row, columnIndexes.get('onlyClub')),
      `${label}のclubのみ`,
    );
    const delivery = {};
    for (const column of holiday.columns) {
      delivery[column.key] = readExcelString(
        row,
        columnIndexes.get(column.key),
        `${label}の${column.header}`,
      );
    }
    delivery.hasDelay = delivery.beforeHolidayDeadline !== '-'
      || delivery.afterHolidayDeadline !== '-';
    products.push({
      name,
      onlyClub,
      category: metadata.category,
      [holiday.key]: delivery,
    });
  }

  return products;
}

async function convertWorkbooks() {
  const results = [];
  for (const holiday of HOLIDAYS) {
    results.push({
      holiday,
      products: await readHolidayWorkbook(holiday),
    });
  }

  for (const { holiday, products } of results) {
    const outputPath = path.join(ROOT, `products-${holiday.key}.json`);
    fs.writeFileSync(outputPath, `${JSON.stringify(products, null, 2)}\n`, 'utf8');
    console.log(`出力: ${path.relative(ROOT, outputPath)} (${products.length} 商品)`);
  }
}

async function main() {
  const command = process.argv[2];
  if (command === 'templates') {
    await createTemplates({ refresh: process.argv.includes('--refresh') });
  } else if (command === 'convert') {
    await convertWorkbooks();
  } else {
    throw new Error('使用方法: node tools/holiday-data.js <templates|convert>');
  }
}

main().catch(error => {
  console.error(`エラー: ${error.message}`);
  process.exitCode = 1;
});
