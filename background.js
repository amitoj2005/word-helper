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
