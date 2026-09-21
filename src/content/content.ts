import { scanPage, preview, fill, debugDom, installHelpers } from './cyberpass';
import type { Message } from '../shared/messages';
chrome.runtime.onMessage.addListener((message: Message, _sender: any, sendResponse: any) => {
  try {
    if (message.type === 'SCAN_PAGE') sendResponse(scanPage());
    else if (message.type === 'PREVIEW') sendResponse(preview(message.requirements));
    else if (message.type === 'FILL') sendResponse(fill(message.requirements, message.options));
    else if (message.type === 'DEBUG_DOM') sendResponse(debugDom());
    else if (message.type === 'INSTALL_HELPERS') { installHelpers(message.requirements); sendResponse({ok:true}); }
  } catch (error) { sendResponse({error: String(error)}); }
  return true;
});
