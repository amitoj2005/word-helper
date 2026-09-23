// Captures the visible tab as a JPEG data URL so content scripts can use it
// as a WebGL texture for the liquid-glass background refraction effect.
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type !== 'CAPTURE_SCREEN' && msg.type !== 'CAPTURE_SCREEN_HD') return false;
  const windowId = sender?.tab?.windowId;
  if (!windowId) {
    console.error('[Word Helper BG] no sender.tab.windowId');
    sendResponse({ dataUrl: null });
    return true;
  }
  // HD variant: quality 100 JPEG — sharper source for the HiDPI WebGL canvas
  const opts = msg.type === 'CAPTURE_SCREEN_HD'
    ? { format: 'jpeg', quality: 100 }
    : { format: 'jpeg', quality: 88 };
  chrome.tabs.captureVisibleTab(windowId, opts, (dataUrl) => {
    if (chrome.runtime.lastError) {
      console.error('[Word Helper BG] captureVisibleTab error:', chrome.runtime.lastError.message);
      sendResponse({ dataUrl: null });
    } else {
      console.log('[Word Helper BG] captured', dataUrl?.length, 'chars');
      sendResponse({ dataUrl });
    }
  });
  return true;
});

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: 'wh-lookup',
    title: 'Look up "%s"',
    contexts: ['selection'],
    documentUrlPatterns: ['https://docs.google.com/document/*']
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId !== 'wh-lookup') return;
  const word = info.selectionText?.trim();
  if (!word) return;
  chrome.tabs.sendMessage(
    tab.id,
    { type: 'WH_LOOKUP', word },
    { frameId: info.frameId }
  );
});
