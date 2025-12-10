// 秋月電子 領収書エクスポーター - Popup Script

document.addEventListener('DOMContentLoaded', () => {
  const downloadBtn = document.getElementById('download-btn');
  const filenamePatternInput = document.getElementById('filename-pattern');
  const downloadReceiptCheckbox = document.getElementById('download-receipt');
  const downloadDeliverySlipCheckbox = document.getElementById('download-delivery-slip');
  const downloadInvoiceCheckbox = document.getElementById('download-invoice');
  const statusDiv = document.getElementById('status');
  const optionsLink = document.getElementById('options-link');

  // 設定をロード
  chrome.storage.sync.get([
    'filenamePattern', 'downloadReceipt', 'downloadDeliverySlip', 'downloadInvoice'
  ], (result) => {
    if (result.filenamePattern) filenamePatternInput.value = result.filenamePattern;
    downloadReceiptCheckbox.checked = result.downloadReceipt !== false;
    downloadDeliverySlipCheckbox.checked = result.downloadDeliverySlip !== false;
    downloadInvoiceCheckbox.checked = result.downloadInvoice !== false;
  });

  // 設定を保存
  filenamePatternInput.addEventListener('change', () => {
    chrome.storage.sync.set({ filenamePattern: filenamePatternInput.value });
  });
  downloadReceiptCheckbox.addEventListener('change', () => {
    chrome.storage.sync.set({ downloadReceipt: downloadReceiptCheckbox.checked });
  });
  downloadDeliverySlipCheckbox.addEventListener('change', () => {
    chrome.storage.sync.set({ downloadDeliverySlip: downloadDeliverySlipCheckbox.checked });
  });
  downloadInvoiceCheckbox.addEventListener('change', () => {
    chrome.storage.sync.set({ downloadInvoice: downloadInvoiceCheckbox.checked });
  });

  // ダウンロードボタン
  downloadBtn.addEventListener('click', async () => {
    if (!downloadReceiptCheckbox.checked && !downloadDeliverySlipCheckbox.checked && !downloadInvoiceCheckbox.checked) {
      setStatus('⚠️ ダウンロードする書類を選択してください', 'error');
      return;
    }
    
    downloadBtn.disabled = true;
    setStatus('📋 注文情報を取得中...', '');
    
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    
    if (!tab?.url?.includes('akizukidenshi.com/catalog/customer/history')) {
      setStatus('⚠️ 秋月電子の注文履歴ページを開いてください', 'error');
      downloadBtn.disabled = false;
      return;
    }

    chrome.tabs.sendMessage(tab.id, { action: 'get_selected_orders' }, (response) => {
      if (chrome.runtime.lastError) {
        setStatus('❌ エラー: ページをリロードしてください', 'error');
        downloadBtn.disabled = false;
        return;
      }

      if (!response?.orders?.length) {
        setStatus('⚠️ 注文が選択されていません', 'error');
        downloadBtn.disabled = false;
        return;
      }

      // 宛名チェック
      const ordersWithoutAddr = response.orders.filter(o => !o.addrComp && !o.addrDept && !o.addrName);
      if (ordersWithoutAddr.length > 0) {
        setStatus('⚠️ 宛名（会社名/部署名/氏名）が未入力の注文があります', 'error');
        downloadBtn.disabled = false;
        return;
      }

      setStatus(`📥 ${response.orders.length}件の注文を処理します...`, '');
      setStatus('⏳ ページがリロードされながら処理が進みます', '');

      // ダウンロードキュー開始
      chrome.tabs.sendMessage(tab.id, {
        action: 'start_download_queue',
        orders: response.orders,
        settings: {
          pattern: filenamePatternInput.value,
          downloadReceipt: downloadReceiptCheckbox.checked,
          downloadDeliverySlip: downloadDeliverySlipCheckbox.checked,
          downloadInvoice: downloadInvoiceCheckbox.checked
        }
      });
      
      // ポップアップを閉じる（ページリロードのため）
      setTimeout(() => window.close(), 1000);
    });
  });

  // ステータス更新
  chrome.runtime.onMessage.addListener((message) => {
    if (message.action === 'update_status') {
      const isComplete = message.status.includes('完了');
      setStatus(isComplete ? `✅ ${message.status}` : message.status, isComplete ? 'success' : '');
      if (isComplete) downloadBtn.disabled = false;
    }
  });

  optionsLink.addEventListener('click', (e) => {
    e.preventDefault();
    chrome.runtime.openOptionsPage();
  });

  function setStatus(text, type) {
    statusDiv.textContent = text;
    statusDiv.className = 'status' + (type ? ' ' + type : '');
  }
});
