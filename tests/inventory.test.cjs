const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const context = vm.createContext({});
vm.runInContext(fs.readFileSync(require('node:path').join(__dirname, '../gas/03_DriveInventory.js'), 'utf8'), context);

function iterator(count, folder = false) {
  let index = 0;
  return {
    hasNext: () => index < count,
    next: () => {
      const id = String(++index);
      return {
        getId: () => id, getName: () => '資料' + id,
        ...(folder ? {} : {getMimeType: () => 'text/plain'}),
        getUrl: () => 'https://drive.google.com/file/d/' + id,
        getLastUpdated: () => new Date('2026-10-15T00:00:00Z')
      };
    }
  };
}

test('上限を超えた資料は切り詰めを明示する', () => {
  const result = context.collectInventory_(iterator(3), 2);
  assert.equal(result.items.length, 2);
  assert.equal(result.truncated, true);
  assert.equal(result.items[0].updatedAt, '2026-10-15T00:00:00.000Z');
});
test('ちょうど上限・空フォルダでは切り詰めにしない', () => {
  assert.equal(context.collectInventory_(iterator(2), 2).truncated, false);
  assert.equal(context.collectInventory_(iterator(0), 2).items.length, 0);
});
test('フォルダの種類を表現できる', () => {
  assert.equal(context.collectInventory_(iterator(1, true), 2).items[0].mimeType, 'application/vnd.google-apps.folder');
});
test('無効な上限は拒否する', () => {
  assert.throws(() => context.collectInventory_(iterator(1), 0));
});
