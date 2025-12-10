// 秋月電子 領収書エクスポーター - Content Script

let spreadsheetData = null;
let spreadsheetSettings = null;
let autoFillSettings = null;
let isInitialized = false;

async function init() {
  const isOrderPage = window.location.href.includes('history.aspx');
  if (!isOrderPage) return;

  await loadSettings();
  isInitialized = true;
  addCheckboxes();
  
}

async function loadSettings() {
  return new Promise((resolve) => {
    chrome.storage.sync.get([
      'spreadsheetUrl', 'sheetName', 'columnName', 'headerRow',
      'autoFillCompany', 'autoFillDepartment'
    ], async (result) => {
      spreadsheetSettings = result;
      autoFillSettings = {
        company: result.autoFillCompany || '',
        department: result.autoFillDepartment || ''
      };
      
      if (result.spreadsheetUrl) {
        chrome.storage.local.get(['spreadsheetData', 'spreadsheetDataTime'], async (cache) => {
          const cacheAge = Date.now() - (cache.spreadsheetDataTime || 0);
          if (cache.spreadsheetData && cacheAge < 300000) {
            spreadsheetData = cache.spreadsheetData;
            resolve();
          } else {
            await fetchSpreadsheetData();
            resolve();
          }
        });
      } else {
        resolve();
      }
    });
  });
}

async function fetchSpreadsheetData() {
  if (!spreadsheetSettings?.spreadsheetUrl) return;
  
  try {
    const match = spreadsheetSettings.spreadsheetUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) return;
    
    const csvUrl = `https://docs.google.com/spreadsheets/d/${match[1]}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(spreadsheetSettings.sheetName || 'シート1')}`;
    const response = await fetch(csvUrl);
    if (!response.ok) return;
    
    const rows = parseCSV(await response.text());
    const columnIndex = columnLetterToIndex(spreadsheetSettings.columnName || 'A');
    const headerRow = (spreadsheetSettings.headerRow || 2) - 1;
    
    spreadsheetData = [];
    for (let i = headerRow; i < rows.length; i++) {
      if (rows[i]?.[columnIndex]?.trim()) {
        spreadsheetData.push(rows[i][columnIndex].trim());
      }
    }
    
    chrome.storage.local.set({ spreadsheetData, spreadsheetDataTime: Date.now() });
  } catch (e) {
    spreadsheetData = null;
  }
}

function parseCSV(csvText) {
  const rows = [];
  let currentRow = [], currentCell = '', inQuotes = false;
  
  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i], nextChar = csvText[i + 1];
    
    if (inQuotes) {
      if (char === '"' && nextChar === '"') { currentCell += '"'; i++; }
      else if (char === '"') inQuotes = false;
      else currentCell += char;
    } else {
      if (char === '"') inQuotes = true;
      else if (char === ',') { currentRow.push(currentCell); currentCell = ''; }
      else if (char === '\n' || (char === '\r' && nextChar === '\n')) {
        currentRow.push(currentCell);
        rows.push(currentRow);
        currentRow = []; currentCell = '';
        if (char === '\r') i++;
      } else if (char !== '\r') currentCell += char;
    }
  }
  if (currentCell || currentRow.length) { currentRow.push(currentCell); rows.push(currentRow); }
  return rows;
}

function columnLetterToIndex(letter) {
  letter = letter.toUpperCase();
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

function addCheckboxes() {
  if (!isInitialized) return;
  
  document.querySelectorAll('.block-purchase-history--frame').forEach((frame) => {
    if (frame.querySelector('.akizuki-exporter-container') || frame.dataset.akizukiExporterProcessed) return;
    if (frame.querySelector('.block-purchase-history--order-cancel-data')) {
      frame.dataset.akizukiExporterProcessed = 'true';
      return;
    }

    const container = document.createElement('div');
    container.className = 'akizuki-exporter-container';
    
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'akizuki-exporter-checkbox';
    
    // baseOrderIdをデータ属性に保存
    const baseOrderId = frame.querySelector('input[name="base_order_id"]')?.value || '';
    checkbox.dataset.baseOrderId = baseOrderId;
    
    checkbox.addEventListener('click', (e) => e.stopPropagation());
    checkbox.addEventListener('change', (e) => {
      const dropdown = container.querySelector('.akizuki-exporter-dropdown');
      if (dropdown) dropdown.disabled = !e.target.checked;
      if (e.target.checked && autoFillSettings) autoFillAddressFields(frame);
    });
    container.appendChild(checkbox);
    
    const label = document.createElement('span');
    label.className = 'akizuki-exporter-label';
    label.textContent = '選択';
    container.appendChild(label);
    
    if (spreadsheetData?.length > 0) {
      const divider = document.createElement('span');
      divider.className = 'akizuki-exporter-divider';
      divider.textContent = '|';
      container.appendChild(divider);
      
      const dropLabel = document.createElement('span');
      dropLabel.className = 'akizuki-exporter-label';
      dropLabel.textContent = '分類:';
      container.appendChild(dropLabel);
      
      const dropdown = document.createElement('select');
      dropdown.className = 'akizuki-exporter-dropdown';
      dropdown.disabled = true;
      dropdown.innerHTML = '<option value="">-- 選択 --</option>' + 
        spreadsheetData.map(item => `<option value="${item}">${item.length > 30 ? item.substring(0, 30) + '...' : item}</option>`).join('');
      dropdown.addEventListener('click', (e) => e.stopPropagation());
      container.appendChild(dropdown);
    }

    const firstTd = frame.querySelector('tbody tr td');
    if (firstTd) firstTd.insertBefore(container, firstTd.firstChild);
    frame.dataset.akizukiExporterProcessed = 'true';
  });
}

function autoFillAddressFields(frame) {
  if (!autoFillSettings) return;
  
  const companyInput = frame.querySelector('input[name="addr_comp"]');
  const deptInput = frame.querySelector('input[name="addr_dept"]');
  
  if (companyInput && autoFillSettings.company) {
    companyInput.value = autoFillSettings.company;
  }
  if (deptInput && autoFillSettings.department) {
    deptInput.value = autoFillSettings.department;
  }
}

// 注文フレームをbaseOrderIdで検索
function findFrameByBaseOrderId(baseOrderId) {
  const frames = document.querySelectorAll('.block-purchase-history--frame');
  for (const frame of frames) {
    const input = frame.querySelector('input[name="base_order_id"]');
    if (input && input.value === baseOrderId) {
      return frame;
    }
  }
  return null;
}

// 選択された注文情報を取得
function getSelectedOrders() {
  const selectedOrders = [];
  
  document.querySelectorAll('.block-purchase-history--frame').forEach((frame) => {
    const checkbox = frame.querySelector('.akizuki-exporter-checkbox');
    if (!checkbox?.checked) return;
    
    const dropdown = frame.querySelector('.akizuki-exporter-dropdown');
    const orderInfo = extractOrderInfo(frame, dropdown?.value || '');
    if (orderInfo) {
      selectedOrders.push(orderInfo);
    }
  });
  
  return selectedOrders;
}

function extractOrderInfo(frame, dropdownValue) {
  try {
    const orderDtEl = frame.querySelector('.block-purchase-history--order_dt');
    const dateMatch = orderDtEl?.textContent.match(/(\d{4}\/\d{1,2}\/\d{1,2})/);
    const dateText = dateMatch ? dateMatch[1] : '';

    const totalEl = frame.querySelector('.block-purchase-history--total');
    const priceMatch = totalEl?.textContent.match(/￥([\d,]+)/);
    const price = priceMatch ? '￥' + priceMatch[1] : '';

    const baseOrderId = frame.querySelector('input[name="base_order_id"]')?.value || '';
    const orderIdLink = frame.querySelector('.block-purchase-history--order-detail a[href*="historydetail"]');
    const orderId = orderIdLink?.textContent.trim() || '';

    const productNameEl = frame.querySelector('.block-purchase-history--goods-name');
    let productName = productNameEl?.textContent.trim() || '';
    if (productName.length > 50) productName = productName.substring(0, 50) + '...';

    // 宛名情報
    const form = frame.querySelector('form[name="frmAddress"]');
    const addrComp = form?.querySelector('input[name="addr_comp"]')?.value || '';
    const addrDept = form?.querySelector('input[name="addr_dept"]')?.value || '';
    const addrName = form?.querySelector('input[name="addr_name"]')?.value || '';
    const csrfToken = form?.querySelector('input[name="crsirefo_hidden"]')?.value || '';

    // hiddenだがchecked属性が入っているケースに対応し、空でなければtrue扱い
    const addrCompOn = form?.querySelector('input[name="addr_comp_on"]')?.value || '';
    const addrDeptOn = form?.querySelector('input[name="addr_dept_on"]')?.value || '';
    const addrNameOn = form?.querySelector('input[name="addr_name_on"]')?.value || '';

    return {
      date: dateText,
      baseOrderId,
      orderId,
      price,
      productName,
      dropdownValue,
      addrComp,
      addrDept,
      addrName,
      addrCompOn,
      addrDeptOn,
      addrNameOn,
      csrfToken
    };
  } catch (e) {
    return null;
  }
}

// ===== ダウンロードタスク管理 =====

// ダウンロードキューを保存して処理開始
async function startDownloadQueue(orders, settings) {
  const queue = [];
  
  for (const order of orders) {
    if (settings.downloadReceipt && order.baseOrderId) {
      queue.push({ ...order, docType: '領収書', type: 'receipt' });
    }
    if (settings.downloadDeliverySlip && order.orderId) {
      queue.push({ ...order, docType: '納品書', type: 'delivery' });
    }
    if (settings.downloadInvoice && order.orderId) {
      queue.push({ ...order, docType: '請求書', type: 'invoice' });
    }
  }
  
  if (queue.length === 0) {
    alert('ダウンロード対象がありません');
    return;
  }
  
  await chrome.storage.local.set({
    downloadQueue: queue,
    downloadIndex: 0,
    downloadTotal: queue.length,
    downloadSuccess: 0,
    downloadFail: 0
  });
  processNextInQueue();
}

// キューから次のアイテムを処理
async function processNextInQueue() {
  const data = await chrome.storage.local.get([
    'downloadQueue', 'downloadIndex', 'downloadTotal', 'downloadSuccess', 'downloadFail'
  ]);
  
  const queue = data.downloadQueue || [];
  const index = data.downloadIndex || 0;
  const total = data.downloadTotal || 0;
  
  if (index >= queue.length) {
    // 完了
    const finalStatus = `完了: ${data.downloadSuccess || 0}件 (失敗: ${data.downloadFail || 0}件)`;
    showNotification(finalStatus);
    chrome.runtime.sendMessage({ action: 'update_status', status: finalStatus });
    await chrome.storage.local.remove(['downloadQueue', 'downloadIndex', 'downloadTotal', 'downloadSuccess', 'downloadFail', 'downloadSettings', 'currentItem']);
    return;
  }
  
  const item = queue[index];
  showProgress(index, total, `処理中... (${index + 1}/${total})`, `${item.date} - ${item.orderId || item.baseOrderId}`);
  
  // 現在処理中のアイテムを保存
  await chrome.storage.local.set({ currentItem: item });
  
  // 宛名フラグを付けてURL生成
  const flags = [];
  if (item.addrComp || item.addrCompOn) flags.push('comp_on=true');
  if (item.addrDept || item.addrDeptOn) flags.push('dept_on=true');
  if (item.addrName || item.addrNameOn) flags.push('name_on=true');
  const flagQuery = flags.length ? '&' + flags.join('&') : '';

  let url = '';
  if (item.type === 'receipt') {
    url = `https://akizukidenshi.com/catalog/customer/printreceipt.aspx?base_order_id=${item.baseOrderId}${flagQuery}`;
  } else if (item.type === 'delivery') {
    url = `https://akizukidenshi.com/catalog/customer/printdeliveryslip.aspx?order_id=${item.orderId}${flagQuery}`;
  } else if (item.type === 'invoice') {
    url = `https://akizukidenshi.com/catalog/customer/printinvoice.aspx?order_id=${item.orderId}${flagQuery}`;
  }

  // 現在処理中のアイテムを保存
  await chrome.storage.local.set({ currentItem: item });

  // サーバ返却PDFを直接ダウンロード
  chrome.runtime.sendMessage({
    action: 'download_binary',
    url: url,
    order: item,
    docType: item.docType
  });
}

// ダウンロード完了時の処理
async function onDownloadComplete(success) {
  const data = await chrome.storage.local.get(['downloadIndex', 'downloadSuccess', 'downloadFail']);
  
  if (success) {
    await chrome.storage.local.set({ downloadSuccess: (data.downloadSuccess || 0) + 1 });
  } else {
    await chrome.storage.local.set({ downloadFail: (data.downloadFail || 0) + 1 });
  }
  
  // currentItemをクリアしてインデックスを進める
  await chrome.storage.local.set({ 
    currentItem: null,
    downloadIndex: (data.downloadIndex || 0) + 1 
  });
  
  // 少し待ってから次へ
  await new Promise(r => setTimeout(r, 1000));
  processNextInQueue();
}

async function incrementFailAndContinue() {
  const data = await chrome.storage.local.get(['downloadIndex', 'downloadFail']);
  await chrome.storage.local.set({ 
    downloadFail: (data.downloadFail || 0) + 1,
    currentItem: null,
    downloadIndex: (data.downloadIndex || 0) + 1 
  });
  await new Promise(r => setTimeout(r, 500));
  processNextInQueue();
}

// ===== UI =====

function showProgress(current, total, status, detail) {
  let progress = document.getElementById('akizuki-exporter-progress');
  if (!progress) {
    progress = document.createElement('div');
    progress.className = 'akizuki-exporter-progress';
    progress.id = 'akizuki-exporter-progress';
    progress.innerHTML = `
      <div class="akizuki-exporter-progress-header">
        <span class="akizuki-exporter-progress-icon">📄</span>
        <span class="akizuki-exporter-progress-title">秋月電子 領収書エクスポーター</span>
      </div>
      <div class="akizuki-exporter-progress-status"></div>
      <div class="akizuki-exporter-progress-bar-bg">
        <div class="akizuki-exporter-progress-bar" style="width: 0%"></div>
      </div>
      <div class="akizuki-exporter-progress-detail"></div>
    `;
    document.body.appendChild(progress);
  }
  
  const percent = total > 0 ? Math.round((current / total) * 100) : 0;
  progress.querySelector('.akizuki-exporter-progress-status').textContent = status;
  progress.querySelector('.akizuki-exporter-progress-bar').style.width = percent + '%';
  progress.querySelector('.akizuki-exporter-progress-detail').textContent = detail || '';
}

function showNotification(message) {
  let progress = document.getElementById('akizuki-exporter-progress');
  if (progress) {
    progress.querySelector('.akizuki-exporter-progress-status').textContent = message;
    progress.querySelector('.akizuki-exporter-progress-bar').style.width = '100%';
    progress.classList.add('success');
    setTimeout(() => progress.remove(), 5000);
  }
}

// ===== メッセージリスナー =====

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'get_selected_orders') {
    sendResponse({ orders: getSelectedOrders() });
    return true;
  }
  
  if (request.action === 'start_download_queue') {
    startDownloadQueue(request.orders, request.settings);
    return true;
  }
  
  if (request.action === 'download_item_complete') {
    onDownloadComplete(request.success);
    return true;
  }
  
  if (request.action === 'refresh_spreadsheet') {
    (async () => {
      await fetchSpreadsheetData();
      document.querySelectorAll('.block-purchase-history--frame').forEach(el => {
        delete el.dataset.akizukiExporterProcessed;
        el.querySelector('.akizuki-exporter-container')?.remove();
      });
      addCheckboxes();
      sendResponse({ success: true });
    })();
    return true;
  }
  
  if (request.action === 'reload_settings') {
    loadSettings().then(() => sendResponse({ success: true }));
    return true;
  }
});

const observer = new MutationObserver(() => { if (isInitialized) addCheckboxes(); });
observer.observe(document.body, { childList: true, subtree: true });

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
