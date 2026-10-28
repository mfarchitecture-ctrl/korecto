const MENU = 'korecto-selection';
function createMenu() {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({ id: MENU, title: 'Corriger avec Korecto', contexts: ['selection'] });
  });
}
chrome.runtime.onInstalled.addListener(createMenu);
chrome.runtime.onStartup.addListener(createMenu);
chrome.contextMenus.onClicked.addListener((info) => {
  if (info.menuItemId !== MENU) return;
  const text = info.selectionText;
  if (typeof text !== 'string' || !text.trim() || text.length > 30000) {
    showError('Sélectionnez entre 1 et 30 000 caractères.'); return;
  }
  chrome.runtime.sendNativeMessage('com.korecto.selection', { action: 'selection', text }, (reply) => {
    const error = chrome.runtime.lastError;
    if (error || !reply || !reply.ok) {
      showError(reply?.error || 'Liaison locale indisponible. Lancez ACTIVER_EXTENSION.cmd dans le dossier de test.');
    } else {
      chrome.action.setBadgeText({ text: '' });
      chrome.action.setTitle({ title: 'Sélection ouverte dans Korecto' });
    }
  });
});
function showError(message) {
  chrome.action.setBadgeText({ text: '!' });
  chrome.action.setBadgeBackgroundColor({ color: '#a04432' });
  chrome.action.setTitle({ title: message });
}
