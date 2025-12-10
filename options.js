// 秋月電子 領収書エクスポーター - Options Script

document.addEventListener('DOMContentLoaded', () => {
  const filenamePatternInput = document.getElementById('filename-pattern');
  const downloadReceiptCheckbox = document.getElementById('download-receipt');
  const downloadDeliverySlipCheckbox = document.getElementById('download-delivery-slip');
  const downloadInvoiceCheckbox = document.getElementById('download-invoice');
  const autoFillCompanyInput = document.getElementById('auto-fill-company');
  const autoFillDepartmentInput = document.getElementById('auto-fill-department');
  const spreadsheetUrlInput = document.getElementById('spreadsheet-url');
  const sheetNameInput = document.getElementById('sheet-name');
  const columnNameInput = document.getElementById('column-name');
  const headerRowInput = document.getElementById('header-row');
  const testFetchBtn = document.getElementById('test-fetch-btn');
  const testResult = document.getElementById('test-result');
  const saveBtn = document.getElementById('save-btn');
  const statusDiv = document.getElementById('status');

  // 設定をロード
  chrome.storage.sync.get([
    'filenamePattern',
    'downloadReceipt',
    'downloadDeliverySlip',
    'downloadInvoice',
    'autoFillCompany',
    'autoFillDepartment',
    'spreadsheetUrl',
    'sheetName',
    'columnName',
    'headerRow'
  ], (result) => {
    if (result.filenamePattern) filenamePatternInput.value = result.filenamePattern;
    // デフォルトはtrue（チェック済み）
    downloadReceiptCheckbox.checked = result.downloadReceipt !== false;
    downloadDeliverySlipCheckbox.checked = result.downloadDeliverySlip !== false;
    downloadInvoiceCheckbox.checked = result.downloadInvoice !== false;
    if (result.autoFillCompany) autoFillCompanyInput.value = result.autoFillCompany;
    if (result.autoFillDepartment) autoFillDepartmentInput.value = result.autoFillDepartment;
    if (result.spreadsheetUrl) spreadsheetUrlInput.value = result.spreadsheetUrl;
    if (result.sheetName) sheetNameInput.value = result.sheetName;
    if (result.columnName) columnNameInput.value = result.columnName;
    if (result.headerRow) headerRowInput.value = result.headerRow;
  });

  // データをテスト取得
  testFetchBtn.addEventListener('click', async () => {
    const url = spreadsheetUrlInput.value;
    const sheet = sheetNameInput.value.trim();
    const column = columnNameInput.value || 'A';
    const headerRow = parseInt(headerRowInput.value) || 2;

    if (!url) {
      testResult.textContent = '⚠️ スプレッドシートURLを入力してください';
      return;
    }
    
    if (!sheet) {
      testResult.textContent = '⚠️ シート名を入力してください';
      return;
    }

    testFetchBtn.disabled = true;
    testFetchBtn.innerHTML = '<span class="loading"></span>取得中...';
    testResult.textContent = '';

    try {
      const data = await fetchColumnData(url, sheet, column, headerRow);
      if (data && data.length > 0) {
        testResult.innerHTML = `<strong style="color:#006600;">✅ ${data.length}件のデータを取得しました:</strong><br>${data.slice(0, 5).join(', ')}${data.length > 5 ? '...' : ''}`;
      } else {
        testResult.textContent = '⚠️ データが見つかりませんでした（シート名・列・開始行を確認してください）';
      }
    } catch (e) {
      testResult.textContent = '❌ 取得に失敗しました: ' + e.message;
    }

    testFetchBtn.disabled = false;
    testFetchBtn.textContent = 'データをテスト取得';
  });

  // 列データを取得
  async function fetchColumnData(spreadsheetUrl, sheetName, columnName, headerRow) {
    const match = spreadsheetUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) throw new Error('URLが正しくありません');
    
    const spreadsheetId = match[1];
    const sheet = encodeURIComponent(sheetName);
    
    const csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${sheet}`;
    
    const response = await fetch(csvUrl);
    if (!response.ok) throw new Error('取得失敗 - シート名と共有設定を確認してください');
    
    const csvText = await response.text();
    const rows = parseCSV(csvText);
    
    const columnIndex = columnLetterToIndex(columnName || 'A');
    const startRow = (headerRow || 2) - 1;
    
    const data = [];
    for (let i = startRow; i < rows.length; i++) {
      if (rows[i][columnIndex] && rows[i][columnIndex].trim()) {
        data.push(rows[i][columnIndex].trim());
      }
    }
    
    return data;
  }

  function parseCSV(csvText) {
    const rows = [];
    let currentRow = [];
    let currentCell = '';
    let inQuotes = false;
    
    for (let i = 0; i < csvText.length; i++) {
      const char = csvText[i];
      const nextChar = csvText[i + 1];
      
      if (inQuotes) {
        if (char === '"' && nextChar === '"') {
          currentCell += '"';
          i++;
        } else if (char === '"') {
          inQuotes = false;
        } else {
          currentCell += char;
        }
      } else {
        if (char === '"') {
          inQuotes = true;
        } else if (char === ',') {
          currentRow.push(currentCell);
          currentCell = '';
        } else if (char === '\n' || (char === '\r' && nextChar === '\n')) {
          currentRow.push(currentCell);
          rows.push(currentRow);
          currentRow = [];
          currentCell = '';
          if (char === '\r') i++;
        } else if (char !== '\r') {
          currentCell += char;
        }
      }
    }
    
    if (currentCell || currentRow.length > 0) {
      currentRow.push(currentCell);
      rows.push(currentRow);
    }
    
    return rows;
  }

  function columnLetterToIndex(letter) {
    letter = letter.toUpperCase();
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
      index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return index - 1;
  }

  // 保存
  saveBtn.addEventListener('click', () => {
    const sheetName = sheetNameInput.value.trim();
    
    if (spreadsheetUrlInput.value && !sheetName) {
      showStatus('⚠️ シート名を入力してください', 'error');
      return;
    }
    
    // 少なくとも1つは選択されている必要がある
    if (!downloadReceiptCheckbox.checked && !downloadDeliverySlipCheckbox.checked && !downloadInvoiceCheckbox.checked) {
      showStatus('⚠️ 領収書・納品書・請求書のいずれかを選択してください', 'error');
      return;
    }
    
    const settings = {
      filenamePattern: filenamePatternInput.value,
      downloadReceipt: downloadReceiptCheckbox.checked,
      downloadDeliverySlip: downloadDeliverySlipCheckbox.checked,
      downloadInvoice: downloadInvoiceCheckbox.checked,
      autoFillCompany: autoFillCompanyInput.value,
      autoFillDepartment: autoFillDepartmentInput.value,
      spreadsheetUrl: spreadsheetUrlInput.value,
      sheetName: sheetName,
      columnName: columnNameInput.value || 'A',
      headerRow: parseInt(headerRowInput.value) || 2
    };

    chrome.storage.sync.set(settings, () => {
      if (chrome.runtime.lastError) {
        showStatus('❌ 保存に失敗しました: ' + chrome.runtime.lastError.message, 'error');
      } else {
        showStatus('✅ 設定を保存しました！', 'success');
        chrome.storage.local.remove(['spreadsheetData', 'spreadsheetDataTime']);
        
        // 開いているタブに設定変更を通知
        chrome.tabs.query({ url: 'https://akizukidenshi.com/*' }, (tabs) => {
          tabs.forEach(tab => {
            chrome.tabs.sendMessage(tab.id, { action: 'reload_settings' }).catch(() => {});
          });
        });
      }
    });
  });

  function showStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    
    setTimeout(() => {
      statusDiv.className = 'status';
    }, 3000);
  }
});
