import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';

const source = await readFile(new URL('Code.gs', import.meta.url), 'utf8');
const header = ['確認済み', 'キャラクター名', 'DC名/ワールド名', 'ステータス', '1時間予約', '備考'];

test('version API identifies this fix and matches the frontend expected version', async () => {
  const html = await readFile(new URL('index.html', import.meta.url), 'utf8');
  const expected = html.match(/const EXPECTED_API_VERSION = "([^"]+)"/)[1];
  const c = vm.createContext({ContentService: {
    MimeType: {JSON: 'application/json'},
    createTextOutput: text => ({setMimeType: () => text})
  }});
  vm.runInContext(source, c);
  const result = JSON.parse(c.doGet({parameter: {action: 'version'}}));
  assert.equal(result.ok, true);
  assert.equal(result.version, '2026.10.28.2');
  assert.equal(result.version, expected);
  assert.notEqual(result.version, '2026.10.28.1');
});

function fixture(data = []) {
  const rows = [[], header, ...data];
  const sheet = {
    getLastRow: () => rows.length,
    getRange: (row, column, height, width) => {
      assert.ok(height > 0, 'do not request an empty range');
      return {getValues: () => rows.slice(row - 1, row - 1 + height)
        .map(values => values.slice(column - 1, column - 1 + width))};
    }
  };
  const context = vm.createContext({});
  vm.runInContext(source, context, {filename: 'Code.gs'});
  context.getMemberSheet_ = () => sheet;
  context.getReservationSheet_ = () => ({getLastRow: () => 1});
  return {context, sheet};
}

test('first member on row 3 and following member on row 4 are searchable', () => {
  const {context: c, sheet} = fixture([
    [true, 'First Member', 'Mana/Chocobo', 'ご予約待ち/ご予約中', false, ''],
    [true, 'Next Member', 'Gaia/Bahamut', 'ご予約待ち/ご予約中', true, '']
  ]);
  assert.equal(c.findMember_(sheet, 'first member').row, 3);
  assert.equal(c.findMember_(sheet, 'next member').row, 4);
  assert.equal(c.member_('  FIRST   MEMBER  ').found, true);
  assert.equal(c.member_('First Member').eligible, true);
  assert.equal(c.member_('First Member').oneHourBooking, false);
  assert.equal(c.member_('Next Member').oneHourBooking, true);
});

test('a sheet containing only the first member is read; headers and missing names are excluded', () => {
  const {context: c, sheet} = fixture([
    [false, 'Only Member', 'Meteor/Shinryu', 'ご予約待ち/ご予約中', true, '']
  ]);
  assert.equal(c.findMember_(sheet, 'only member').row, 3);
  assert.equal(c.member_('Only Member').eligible, true);
  assert.equal(c.findMember_(sheet, 'キャラクター名'), null);
  assert.equal(c.member_('Missing Member').found, false);
  const empty = fixture();
  assert.equal(empty.context.member_('Missing Member').found, false);
});

test('row 3 retains eligibility rules and only boolean true enables one-hour booking', () => {
  for (const status of ['招待不要', '登録完了', '招待済み', 'ご予約待ち', 'ご予約中', '']) {
    const {context: c} = fixture([[true, 'First Member', 'Mana/Chocobo', status, false, '']]);
    const result = c.member_('First Member');
    assert.equal(result.found, true);
    assert.equal(result.eligible, false, status);
    assert.equal(result.sourceStatus, status);
  }
  for (const value of [true, false, 'TRUE', '1時間予約', '']) {
    const {context: c} = fixture([[true, 'First Member', 'Mana/Chocobo', 'ご予約待ち/ご予約中', value, '']]);
    assert.equal(c.member_('First Member').oneHourBooking, value === true);
  }
});

test('booking on row 3 rejects an ineligible member before any reservation write', () => {
  const {context: c} = fixture([[true, 'First Member', 'Mana/Chocobo', '招待不要', false, '']]);
  c.validateBooking_ = () => {};
  let released = false;
  c.LockService = {getScriptLock: () => ({waitLock: () => {}, releaseLock: () => {released = true;}})};
  const result = c.book_({name: 'First Member'});
  assert.equal(result.ok, false);
  assert.match(result.message, /今回の予約対象ではありません/);
  assert.equal(released, true);
});
