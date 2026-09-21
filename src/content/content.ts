import { scanPage, preview, fill, debugDom, installHelpers } from './cyberpass';
import type { Message } from '../shared/messages';

function isCyberPassQuestionnairePage(): boolean {
  if (location.origin !== 'https://app.fido.cyber-pass.org') return false;
  if (!/^\/procedures\/[^/]+$/.test(location.pathname)) return false;
  return new URLSearchParams(location.search).get('step') === 'fido_user_authenticator_vendor_questionnaire';
}

function applyDescriptionColorOverride(enabled: boolean): void {
  const styleId = 'cyberpass-importer-description-color';
  document.getElementById(styleId)?.remove();
  if (!enabled) return;
  const style = document.createElement('style');
  style.id = styleId;
  style.textContent = '.input-node-view-builder-description, .input-node-view-builder-description * { color: #000 !important; }';
  document.documentElement.appendChild(style);
}

if (!isCyberPassQuestionnairePage()) {
  // The manifest limits injection to procedure pages; this guard further limits
  // behavior to the vendor-questionnaire step on those pages.
} else {
  chrome.storage.local.get({ overrideRequirementColour: false }, (settings: any) => applyDescriptionColorOverride(Boolean(settings.overrideRequirementColour)));
  chrome.runtime.onMessage.addListener((message: Message, _sender: any, sendResponse: any) => {
  try {
    if (message.type === 'SCAN_PAGE') scanPage().then(sendResponse).catch((error: unknown) => sendResponse({error: String(error)}));
    else if (message.type === 'PREVIEW') preview(message.requirements).then(sendResponse).catch((error: unknown) => sendResponse({error: String(error)}));
    else if (message.type === 'FILL') fill(message.requirements, message.options).then(sendResponse).catch((error: unknown) => sendResponse({error: String(error)}));
    else if (message.type === 'DEBUG_DOM') sendResponse(debugDom());
    else if (message.type === 'INSTALL_HELPERS') { installHelpers(message.requirements); sendResponse({ok:true}); }
    else if (message.type === 'SET_DESCRIPTION_COLOR') { applyDescriptionColorOverride(message.enabled); sendResponse({ok:true}); }
  } catch (error) { sendResponse({error: String(error)}); }
  return true;
  });
}
