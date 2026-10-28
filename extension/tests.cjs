const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const crypto = require('node:crypto');
const path = require('node:path');
const manifest = JSON.parse(fs.readFileSync(path.join(__dirname, 'manifest.json'), 'utf8'));
assert.deepEqual(manifest.permissions, ['contextMenus', 'nativeMessaging']);
assert.equal(manifest.host_permissions, undefined);
let click, installed, menu, sent = [], badge, reply = { ok: true };
const chrome = {
  contextMenus: { removeAll: cb => cb(), create: value => { menu = value; }, onClicked: { addListener: cb => { click = cb; } } },
  runtime: { lastError: null, onInstalled: { addListener: cb => { installed = cb; } }, onStartup: { addListener: () => {} }, sendNativeMessage: (host, message, cb) => { sent.push({ host, message }); cb(reply); } },
  action: { setBadgeText: value => { badge = value.text; }, setBadgeBackgroundColor: () => {}, setTitle: () => {} }
};
vm.runInNewContext(fs.readFileSync(path.join(__dirname, 'background.js'), 'utf8'), { chrome });
installed();
assert.equal(sent.length, 0);
assert.equal(menu.title, 'Corriger avec Korecto');
assert.equal(menu.contexts.join(','), 'selection');
click({ menuItemId: menu.id, selectionText: 'Bonjour 😊\nTexte choisi', pageUrl: 'https://exemple.test/prive' });
assert.equal(sent.length, 1);
assert.equal(sent[0].host, 'com.korecto.selection');
assert.equal(sent[0].message.text, 'Bonjour 😊\nTexte choisi');
assert.deepEqual(Object.keys(sent[0].message).sort(), ['action', 'text']);
assert.equal(badge, '');
for (const value of ['', ' ', 'x'.repeat(30001)]) click({ menuItemId: menu.id, selectionText: value });
assert.equal(sent.length, 1);
reply = { ok: false, error: 'Liaison absente' }; click({ menuItemId: menu.id, selectionText: 'Texte' }); assert.equal(badge, '!');
const id = [...crypto.createHash('sha256').update(Buffer.from(manifest.key, 'base64')).digest().subarray(0, 16)].map(x => String.fromCharCode(97 + (x >> 4), 97 + (x & 15))).join('');
assert.ok(fs.readFileSync(path.join(__dirname, '..', 'src', 'SelectionAccess.cs'), 'utf8').includes(`chrome-extension://${id}/`));
console.log('6 groupes de contrôles extension réussis : permissions, menu sans envoi automatique, texte Unicode seul, limites, erreur, identité hôte.');
