import { scanPage, preview, fill, debugDom, installHelpers } from './cyberpass';
import type { Message } from '../shared/messages';

function isCyberPassQuestionnairePage(): boolean {
  if (location.origin !== 'https://app.fido.cyber-pass.org') return false;
  if (!/^\/procedures\/[^/]+$/.test(location.pathname)) return false;
  return new URLSearchParams(location.search).get('step') === 'fido_user_authenticator_vendor_questionnaire';
}

if (!isCyberPassQuestionnairePage()) {
  // The manifest limits injection to procedure pages; this guard further limits
  // behavior to the vendor-questionnaire step on those pages.
} else chrome.runtime.onMessage.addListener((message: Message, _sender: any, sendResponse: any) => {
  try {
    if (message.type === 'SCAN_PAGE') sendResponse(scanPage());
    else if (message.type === 'PREVIEW') sendResponse(preview(message.requirements));
    else if (message.type === 'FILL') sendResponse(fill(message.requirements, message.options));
    else if (message.type === 'DEBUG_DOM') sendResponse(debugDom());
    else if (message.type === 'INSTALL_HELPERS') { installHelpers(message.requirements); sendResponse({ok:true}); }
  } catch (error) { sendResponse({error: String(error)}); }
  return true;
});
