let products = [];

const categoryNames = {
  print: '紙に印刷する',
  hanko: '印鑑・ハンコ作成',
  naire: '名入れ・ノベルティ',
};

const holidayPageConfig = {
  goldenWeek: {
    label: 'ゴールデンウィーク',
    title: 'ゴールデンウィーク期間の商品納期変更について',
    productsFile: 'products-goldenWeek.json',
    columns: [
      { key: 'normalDelivery', label: 'ゴールデンウィーク期間納期', width: 100 },
      { key: 'deadline', label: '締切時間', width: 100 },
      { key: 'period', label: 'ゴールデンウィーク期間', width: 200 },
      { key: 'beforeHolidayDeadline', label: '期間前<br>締切', width: 100 },
      { key: 'afterHolidayDeadline', label: '期間後<br>締切', width: 100 },
    ],
  },
  obon: {
    label: 'お盆',
    title: 'お盆期間の商品納期変更について',
    productsFile: 'products-obon.json',
    columns: [
      { key: 'normalDelivery', label: 'お盆期間納期', width: 100 },
      { key: 'deadline', label: '締切時間', width: 100 },
      { key: 'period', label: 'お盆期間', width: 200 },
      { key: 'beforeHolidayDeadline', label: '期間前<br>締切', width: 100 },
      { key: 'afterHolidayDeadline', label: '期間後<br>締切', width: 100 },
    ],
  },
  yearEndNewYear: {
    label: '年末年始',
    title: '年末年始期間の商品納期変更について',
    productsFile: 'products-yearEndNewYear.json',
    columns: [
      { key: 'normalDelivery', label: '通常納期', width: 120 },
      { key: 'period', label: '年末年始期間', width: 180 },
      { key: 'beforeHolidayDeadline', label: '年内に届く<br>最終ご注文日', width: 180 },
      { key: 'afterHolidayDeadline', label: '年初すぐに届く<br>ご注文日', width: 180 },
    ],
  },
};

function getHolidayPage() {
  const params = new URLSearchParams(window.location.search);
  const pageFromQuery = params.get('holidayPage');
  const pageFromElement = document.querySelector('[data-holiday-page]')?.dataset?.holidayPage;
  const pageFromGlobal = window.holidayPage;
  const holidayPage = pageFromQuery || pageFromElement || pageFromGlobal || 'obon';
  return holidayPageConfig[holidayPage] ? holidayPage : 'obon';
}

const holidayPage = getHolidayPage();

function updateHolidayTexts() {
  const config = holidayPageConfig[holidayPage];
  document.title = config.title;
  const heading = document.getElementById('pageHeading');
  if (heading) {
    heading.textContent = `パプリからのお知らせ - ${config.label}`;
  }
}

// チャンネル判定.
function getCurrentPlatform() {
  const params = new URLSearchParams(window.location.search);
  const platformParam = params.get('platform');
  if (platformParam && ['club', 'tryon', 'epro'].includes(platformParam)) {
    return platformParam;
  }

  const hostname = window.location.hostname;
  if (hostname.includes('solspc')) {
    return 'tryon';
  } else if (hostname.includes('eprospc')) {
    return 'epro';
  } else if (hostname.includes('spc')) {
    return 'club';
  }
  return 'club';
}

const currentPlatform = getCurrentPlatform();

async function loadProducts() {
  try {
    const productsFile = holidayPageConfig[holidayPage].productsFile;
    const response = await fetch(`/j3a/js/news/common/${productsFile}`);
    if (!response.ok) {
      throw new Error(`${productsFile} の読み込みに失敗しました: ${response.status}`);
    }
    products = await response.json();
  } catch (error) {
    console.error(error);
    products = [];
  }
}

const searchInput = document.getElementById('searchInput');
const categorySelect = document.getElementById('categorySelect');
const productSelect = document.getElementById('productSelect');
const productsContainer = document.getElementById('productsContainer');

let searchTerm = '';
let selectedCategory = 'all';
let selectedProduct = 'all';

// 商品セレクトボックス更新.
function updateProductSelect() {
  productSelect.innerHTML = '<option value="all">すべての商品</option>';

  if (selectedCategory === 'all') {
    productSelect.disabled = true;
    return;
  }

  productSelect.disabled = false;

  const categoryProducts = products.filter(p => p.category === selectedCategory);
  // 商品名重複を除外.
  const uniqueProductNames = [...new Set(categoryProducts.map(p => p.name))];
  uniqueProductNames.forEach(productName => {
    const option = document.createElement('option');
    option.value = productName;
    option.textContent = productName;
    productSelect.appendChild(option);
  });
}

// フィルタリング.
const filterProducts = () => {
  return products.filter(product => {
    const matchesSearch = product.name.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesCategory = selectedCategory === 'all' || product.category === selectedCategory;
    const matchesProduct = selectedProduct === 'all' || product.name === selectedProduct;
    // onlyClub が true の場合、クラブページ以外では表示しない.
    const matchesPlatform = currentPlatform === 'club' || !product.onlyClub;
    return matchesSearch && matchesCategory && matchesProduct && matchesPlatform;
  });
}

// 納期変更のアイコン生成.
// const getDelayBadgeIcon = () => {
//   return `<img src="/j3a/images/news/common/icon-delay.svg" alt="#" />`;
// }
const getDelayBadgeIcon = () => { return ``; }

const getNoDelayBadgeIcon = () => { return `<img src="/j3a/images/news/common/icon-normal.svg" alt="#" />`; }

// テーブル行生成.
function createTableRow(product) {
  const defaultHoliday = {
    period: '-',
    normalDelivery: '-',
    deadline: '-',
    beforeHolidayDeadline: '-',
    afterHolidayDeadline: '-',
    hasDelay: false,
  };
  const holiday = { ...defaultHoliday, ...(product[holidayPage] || {}) };
  const badgeClass = holiday.hasDelay ? 'badge-delay' : 'badge-no-delay';
  const badgeText = holiday.hasDelay ? 'あり' : 'なし';
  const badgeIcon = holiday.hasDelay ? getDelayBadgeIcon() : getNoDelayBadgeIcon();
  const rowClass = holiday.hasDelay ? 'delay' : 'no-delay';
  const holidayCells = holidayPageConfig[holidayPage].columns
    .map(column => `<td>${holiday[column.key]}</td>`)
    .join('');

  return `
    <tr class="${rowClass}">
      <td>${product.name}</td>
      ${holidayCells}
    </tr>
  `;
}

// テーブルヘッダー生成.
function createTableHeader() {
  const columns = holidayPageConfig[holidayPage].columns
    .map(column => `<th style="width: ${column.width}px;">${column.label}</th>`)
    .join('');
  return `
    <thead>
      <tr>
        <th>商品名</th>
        ${columns}
      </tr>
    </thead>
  `;
}

// 商品を表示.
function renderProducts() {
  const filteredProducts = filterProducts();

  if (filteredProducts.length === 0) {
    productsContainer.innerHTML = `
      <div class="table-container">
        <div class="empty-state">該当する商品が見つかりませんでした</div>
      </div>
    `;
    return;
  }

  // カテゴリー別に表示.
  if (selectedCategory === 'all') {
    const groupedProducts = {
      print: filteredProducts.filter(p => p.category === 'print'),
      hanko: filteredProducts.filter(p => p.category === 'hanko'),
      naire: filteredProducts.filter(p => p.category === 'naire'),
    };

    let html = '';

    Object.entries(groupedProducts).forEach(([category, categoryProducts]) => {
      if (categoryProducts.length === 0) return;

      html += `
        <section class="category-section">
          <h2 class="category-title g-hd g-hd-2 g-hd-${category}">${categoryNames[category]}</h2>
          <div class="table-container">
            <div class="table-wrapper">
              <table>
                ${createTableHeader()}
                <tbody>
                  ${categoryProducts.map(createTableRow).join('')}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      `;
    });

    productsContainer.innerHTML = html;
  } else {
    // 特定のカテゴリーが選択.
    const html = `
      <div class="table-container">
        <div class="table-wrapper">
          <table>
            ${createTableHeader()}
            <tbody>
              ${filteredProducts.map(createTableRow).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
    productsContainer.innerHTML = html;
  }
}

searchInput.addEventListener('input', (e) => {
  searchTerm = e.target.value;
  renderProducts();
});

categorySelect.addEventListener('change', (e) => {
  selectedCategory = e.target.value;
  selectedProduct = 'all';
  productSelect.value = 'all';
  updateProductSelect();
  renderProducts();
});

productSelect.addEventListener('change', (e) => {
  selectedProduct = e.target.value;
  renderProducts();
});

// 初期表示
loadProducts().then(() => {
  updateHolidayTexts();
  updateProductSelect();
  renderProducts();
});