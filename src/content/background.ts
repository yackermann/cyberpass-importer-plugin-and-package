const workbookKey = (procedureId: string) => `cyberpassWorkbook:${procedureId}`;

chrome.runtime.onMessage.addListener((message: any, _sender: any, sendResponse: (response: any) => void) => {
  if (!['STORE_WORKBOOK', 'GET_WORKBOOK', 'CLEAR_WORKBOOK'].includes(message?.type)) return undefined;
  const key = workbookKey(String(message.procedureId || ''));
  if (message.type === 'STORE_WORKBOOK') {
    chrome.storage.session.set({ [key]: { fileName: message.fileName, requirements: message.requirements } })
      .then(() => sendResponse({ ok: true })).catch((error: unknown) => sendResponse({ error: String(error) }));
  } else if (message.type === 'GET_WORKBOOK') {
    chrome.storage.session.get(key)
      .then((saved: any) => sendResponse({ workbook: saved[key] || null })).catch((error: unknown) => sendResponse({ error: String(error) }));
  } else {
    chrome.storage.session.remove(key)
      .then(() => sendResponse({ ok: true })).catch((error: unknown) => sendResponse({ error: String(error) }));
  }
  return true;
});

chrome.runtime.onInstalled.addListener(() => {
  console.info('CyberPass Excel Importer ready');
});
