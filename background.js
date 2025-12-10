// 秋月電子 領収書エクスポーター - Background Script

let filenamePattern = '{date}_{order_id}_{type}';
let currentRename = null;

chrome.storage.sync.get(['filenamePattern'], (result) => {
  if (result.filenamePattern) filenamePattern = result.filenamePattern;
});

chrome.storage.onChanged.addListener((changes) => {
  if (changes.filenamePattern) filenamePattern = changes.filenamePattern.newValue;
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'convert_html_to_pdf') {
    currentRename = { order: request.order, docType: request.docType };
    convertHtmlToPdf(request.html, request.order, request.docType, request.baseUrl, sender.tab?.id);
  }
  
  if (request.action === 'capture_url') {
    currentRename = { order: request.order, docType: request.docType };
    captureUrlToPdf(request.url, request.order, request.docType, sender.tab?.id);
  }

  if (request.action === 'download_binary') {
    currentRename = { order: request.order, docType: request.docType };
    downloadBinary(request.url, request.order, request.docType, sender.tab?.id);
  }
});

// HTMLをPDFに変換
async function convertHtmlToPdf(html, order, docType, baseUrl, sourceTabId) {
  let tabId = null;
  
  try {
    // baseタグを追加してリソースの相対パスを解決
    const modifiedHtml = html.replace(
      '<head>',
      `<head><base href="${baseUrl}/">`
    );
    
    // data:URLでHTMLを開く
    const dataUrl = 'data:text/html;charset=utf-8,' + encodeURIComponent(modifiedHtml);
    
    const tab = await chrome.tabs.create({ url: dataUrl, active: false });
    if (!tab) throw new Error('Failed to create tab');
    tabId = tab.id;

    // ページ読み込みを待つ
    await new Promise(r => setTimeout(r, 2500));

    await chrome.debugger.attach({ tabId: tabId }, "1.3");

    await chrome.debugger.sendCommand({ tabId: tabId }, "Emulation.setEmulatedMedia", {
      media: "print"
    });

    await new Promise(r => setTimeout(r, 500));

    const result = await chrome.debugger.sendCommand({ tabId: tabId }, "Page.printToPDF", {
      printBackground: true,
      marginTop: 0.4,
      marginBottom: 0.4,
      marginLeft: 0.4,
      marginRight: 0.4,
      paperWidth: 8.27,
      paperHeight: 11.69,
      preferCSSPageSize: true
    });

    const filename = generateFilename(filenamePattern, order, docType);
    const pdfDataUrl = "data:application/pdf;base64," + result.data;

    await new Promise((resolve, reject) => {
      chrome.downloads.download({
        url: pdfDataUrl,
        filename: 'AkizukiReceipts/' + filename,
        saveAs: false
      }, (downloadId) => {
        if (chrome.runtime.lastError) reject(chrome.runtime.lastError);
        else resolve(downloadId);
      });
    });

    console.log('Download success:', filename);

    if (sourceTabId) {
      chrome.tabs.sendMessage(sourceTabId, { action: 'download_item_complete', success: true }).catch(() => {});
    }

  } catch (err) {
    console.error('Convert error:', err);
    if (sourceTabId) {
      chrome.tabs.sendMessage(sourceTabId, { action: 'download_item_complete', success: false, error: err.message }).catch(() => {});
    }
  } finally {
    if (tabId) {
      try { await chrome.debugger.detach({ tabId: tabId }); } catch(e) {}
      await new Promise(r => setTimeout(r, 300));
      chrome.tabs.remove(tabId).catch(() => {});
    }
  }
}

// URLをPDFに変換（フォールバック用）
async function captureUrlToPdf(url, order, docType, sourceTabId) {
  let tabId = null;
  
  try {
    const tab = await chrome.tabs.create({ url: url, active: false });
    if (!tab) throw new Error('Failed to create tab');
    tabId = tab.id;

    // 納品書・請求書は描画完了がやや遅いので少し長めに待つ
    const waitMs = docType === '領収書' ? 3500 : 5000;
    await new Promise(r => setTimeout(r, waitMs));

    await chrome.debugger.attach({ tabId: tabId }, "1.3");

    await chrome.debugger.sendCommand({ tabId: tabId }, "Emulation.setEmulatedMedia", {
      media: "print"
    });

    await new Promise(r => setTimeout(r, 500));

    const result = await chrome.debugger.sendCommand({ tabId: tabId }, "Page.printToPDF", {
      printBackground: true,
      marginTop: 0.4,
      marginBottom: 0.4,
      marginLeft: 0.4,
      marginRight: 0.4,
      paperWidth: 8.27,
      paperHeight: 11.69,
      preferCSSPageSize: true
    });

    const filename = generateFilename(filenamePattern, order, docType);
    const pdfDataUrl = "data:application/pdf;base64," + result.data;

    await new Promise((resolve, reject) => {
      chrome.downloads.download({
        url: pdfDataUrl,
        filename: 'AkizukiReceipts/' + filename,
        saveAs: false
      }, (downloadId) => {
        if (chrome.runtime.lastError) reject(chrome.runtime.lastError);
        else resolve(downloadId);
      });
    });

    if (sourceTabId) {
      chrome.tabs.sendMessage(sourceTabId, { action: 'download_item_complete', success: true }).catch(() => {});
    }

  } catch (err) {
    console.error('Capture error:', err);
    if (sourceTabId) {
      chrome.tabs.sendMessage(sourceTabId, { action: 'download_item_complete', success: false, error: err.message }).catch(() => {});
    }
  } finally {
    if (tabId) {
      try { await chrome.debugger.detach({ tabId: tabId }); } catch(e) {}
      await new Promise(r => setTimeout(r, 300));
      chrome.tabs.remove(tabId).catch(() => {});
    }
  }
}

// バイナリ(PDF)を直接ダウンロード（納品書・請求書用）
async function downloadBinary(url, order, docType, sourceTabId) {
  try {
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    // 常にユーザ指定パターンで命名（ヘッダは使わない）
    const filename = generateFilename(filenamePattern, order, docType);

    const buf = await res.arrayBuffer();
    const base64Data = arrayBufferToBase64(buf);
    const dataUrl = `data:application/pdf;base64,${base64Data}`;

    await new Promise((resolve, reject) => {
      chrome.downloads.download({
        url: dataUrl,
        filename: 'AkizukiReceipts/' + filename,
        saveAs: false
      }, (downloadId) => {
        if (chrome.runtime.lastError) reject(chrome.runtime.lastError);
        else resolve(downloadId);
      });
    });

    if (sourceTabId) {
      chrome.tabs.sendMessage(sourceTabId, { action: 'download_item_complete', success: true }).catch(() => {});
    }
  } catch (err) {
    console.error('Binary download error:', err);
    if (sourceTabId) {
      chrome.tabs.sendMessage(sourceTabId, { action: 'download_item_complete', success: false, error: err.message }).catch(() => {});
    }
  }
}

// ArrayBuffer → Base64
function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function formatDate(dateString) {
  try {
    const match = dateString?.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
    return dateString?.trim() || '';
  } catch (e) {
    return dateString || '';
  }
}

function sanitizeFilename(name) {
  return name.replace(/[<>:"/\\|?*\r\n]/g, '_').replace(/\.+/g, '.').trim();
}

function generateFilename(pattern, order, type) {
  const dateStr = formatDate(order.date);
  const priceStr = order.price ? order.price.replace(/[^\d]/g, '') : '0';
  const orderIdStr = order.orderId || order.baseOrderId || 'unknown';
  const baseOrderIdStr = order.baseOrderId || (orderIdStr !== 'unknown' ? orderIdStr.replace(/-\d{2}$/, '') : 'unknown');
  let productStr = (order.productName || '').replace(/[<>:"/\\|?*\r\n]/g, '_').substring(0, 30);
  let dropdownStr = (order.dropdownValue || '').replace(/[<>:"/\\|?*\r\n]/g, '_');
  
  let name = pattern
    .replace('{date}', dateStr)
    .replace('{order_id}', orderIdStr)
    .replace('{base_order_id}', baseOrderIdStr)
    .replace('{type}', type)
    .replace('{price}', priceStr)
    .replace('{product}', productStr)
    .replace('{dropdown}', dropdownStr);

  if (!name.toLowerCase().endsWith('.pdf')) name += '.pdf';
  return sanitizeFilename(name);
}

// ダウンロードファイル名を強制的に上書き
chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
  if (currentRename) {
    const forcedName = 'AkizukiReceipts/' + generateFilename(filenamePattern, currentRename.order, currentRename.docType);
    currentRename = null;
    suggest({ filename: forcedName, conflictAction: 'overwrite' });
    return;
  }
  suggest();
});
