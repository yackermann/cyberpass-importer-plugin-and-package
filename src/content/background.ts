function exposeSessionStorageToContentScripts(): void {
  chrome.storage.session.setAccessLevel({ accessLevel: 'TRUSTED_AND_UNTRUSTED_CONTEXTS' }).catch(() => undefined);
}

chrome.runtime.onInstalled.addListener(() => {
  exposeSessionStorageToContentScripts();
  console.info('CyberPass Excel Importer ready');
});
chrome.runtime.onStartup.addListener(exposeSessionStorageToContentScripts);
exposeSessionStorageToContentScripts();
