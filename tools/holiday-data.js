const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');

const ROOT = path.join(__dirname, '..');
const EXCEL_DIR = path.join(ROOT, 'excel');
const CATEGORY_LABELS = {
  print: '紙に印刷する',
  hanko: '印鑑・ハンコ作成',
  naire: '名入れ・ノベルティ',
};
const CATEGORY_KEYS = Object.fromEntries(
  Object.entries(CATEGORY_LABELS).map(([key, label]) => [label, key]),
);
const HOLIDAYS = [
  {
    key: 'obon',
    label: 'お盆',
    file: 'obon.xlsx',
    columns: {
      period: 'お盆期間',
      normalDelivery: 'お盆期間納期',
      deadline: '締切時間',
      beforeHolidayDeadline: '期間前締切',
      afterHolidayDeadline: '期間後締切',
    },
  },
  {
    key: 'goldenWeek',
    label: 'ゴールデンウィーク',
    file: 'goldenWeek.xlsx',
    columns: {
      period: 'ゴールデンウィーク期間',
      normalDelivery: 'ゴールデンウィーク期間納期',
      deadline: '締切時間',
      beforeHolidayDeadline: '期間前締切',
      afterHolidayDeadline: '期間後締切',
    },
  },
  {
    key: 'yearEndNewYear',
    label: '年末年始',
    file: 'yearEndNewYear.xlsx',
    columns: {
      period: '年末年始期間',
      normalDelivery: '通常納期',
      deadline: '締切時間（JSON保持用）',
      beforeHolidayDeadline: '年内最終注文日',
      afterHolidayDeadline: '年初注文日',
    },
  },
];

const commonColumns = [
  { key: 'name', header: '商品名', width: 48 },
  { key: 'category', header: 'カテゴリー', width: 22 },
  { key: 'onlyClub', header: 'クラブページのみ', width: 18 },
  { key: 'baseNormalDelivery', header: '商品基本納期', width: 20 },
  { key: 'baseDeadline', header: '商品基本締切時間', width: 22 },
];

function getColumns(holiday) {
  return [
    ...commonColumns,
    ...Object.entries(holiday.columns).map(([key, header]) => ({
      key,
      header,
      width: 24,
    })),
    { key: 'hasDelay', header: '納期変更あり', width: 18 },
  ];
}

function requireString(value, label) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`${label} は空欄にできません。`);
  }
  return value;
}

function readCategory(value, label) {
  const category = CATEGORY_KEYS[value] ? CATEGORY_KEYS[value] : value;
  if (!Object.hasOwn(CATEGORY_LABELS, category)) {
    throw new Error(`${label} のカテゴリーが不正です: ${value}`);
  }
  return category;
}

function readBoolean(value, label) {
  if (value === true || value === 'はい') return true;
  if (value === false || value === 'いいえ') return false;
  throw new Error(`${label} は「はい」または「いいえ」で入力してください。`);
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

async function createTemplates() {
  const sourcePath = path.join(ROOT, 'products.json');
  const source = JSON.parse(fs.readFileSync(sourcePath, 'utf8'));
  if (!Array.isArray(source) || source.length === 0) {
    throw new Error('products.json に商品データがありません。');
  }
  fs.mkdirSync(EXCEL_DIR, { recursive: true });

  for (const holiday of HOLIDAYS) {
    const outputPath = path.join(EXCEL_DIR, holiday.file);
    if (fs.existsSync(outputPath)) {
      throw new Error(`${outputPath} はすでに存在します。上書きしないため処理を中止しました。`);
    }
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
        category: CATEGORY_LABELS[product.category] || product.category,
        onlyClub: product.onlyClub ? 'はい' : 'いいえ',
        baseNormalDelivery: product.normalDelivery,
        baseDeadline: product.deadline,
        ...delivery,
        hasDelay: delivery.hasDelay ? 'はい' : 'いいえ',
      };
      sheet.addRow(columns.map(column => row[column.key] ?? ''));
      const rowNumber = index + 2;
      sheet.getCell(rowNumber, 3).dataValidation = {
        type: 'list',
        allowBlank: false,
        formulae: ['"はい,いいえ"'],
      };
      sheet.getCell(rowNumber, columns.findIndex(column => column.key === 'hasDelay') + 1)
        .dataValidation = {
          type: 'list',
          allowBlank: false,
          formulae: ['"はい,いいえ"'],
        };
    });

    const guide = workbook.addWorksheet('入力方法');
    guide.columns = [{ width: 110 }];
    guide.addRows([
      [`${holiday.label}の納期データ入力用ファイルです。`],
      ['「商品データ」シートの2行目以降を編集してください。'],
      ['商品名・カテゴリー・クラブページのみの列は、JSON出力に必要なので空欄にしないでください。'],
      ['納期・期間・注文日は文字列で入力してください（例：12/25(金)、3営業日後）。'],
      ['商品基本納期／商品基本締切時間は商品共通の値、休業期間列はこの休業期間用の値です。'],
      ['年末年始の締切時間（JSON保持用）はサイト表示には使いませんが、JSONデータ保持のため変更せず残してください。'],
      ['更新後、プロジェクトのルートで npm run holiday:convert を実行すると3種類のJSONを出力します。'],
    ]);
    guide.getColumn(1).alignment = { wrapText: true, vertical: 'top' };
    guide.getColumn(1).width = 110;
    guide.eachRow(row => { row.height = 30; });

    await workbook.xlsx.writeFile(outputPath);
    console.log(`作成: ${path.relative(ROOT, outputPath)} (${source.length} 商品)`);
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

  const products = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (row.values.slice(1, columns.length + 1).every(value => value == null || value === '')) {
      continue;
    }
    const valueByKey = Object.fromEntries(
      columns.map((column, index) => [column.key, getCellValue(row, index + 1)]),
    );
    const label = `${holiday.file} ${rowNumber}行目`;
    const name = requireString(valueByKey.name, `${label}の商品名`);

    const delivery = {
      period: readExcelString(row, 6, `${label}の${holiday.columns.period}`),
      normalDelivery: readExcelString(row, 7, `${label}の${holiday.columns.normalDelivery}`),
      deadline: readExcelString(row, 8, `${label}の${holiday.columns.deadline}`),
      beforeHolidayDeadline: readExcelString(row, 9, `${label}の${holiday.columns.beforeHolidayDeadline}`),
      afterHolidayDeadline: readExcelString(row, 10, `${label}の${holiday.columns.afterHolidayDeadline}`),
      hasDelay: readBoolean(valueByKey.hasDelay, `${label}の納期変更あり`),
    };
    products.push({
      name,
      onlyClub: readBoolean(valueByKey.onlyClub, `${label}のクラブページのみ`),
      category: readCategory(valueByKey.category, `${label}のカテゴリー`),
      normalDelivery: readExcelString(row, 4, `${label}の商品基本納期`),
      deadline: readExcelString(row, 5, `${label}の商品基本締切時間`),
      [holiday.key]: delivery,
    });
  }

  if (products.length === 0) {
    throw new Error(`${holiday.file} に商品データがありません。`);
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
    await createTemplates();
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
