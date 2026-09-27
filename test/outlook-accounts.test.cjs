/* More than one Microsoft account, part 1 of #38: the account model.
 *
 * `outlookAccounts` in data.json is read, `default` is the sign-in that
 * already exists, and a note's `source_account` stamp names its account
 * (absent means default). Nothing signs in, nothing syncs, nothing shows.
 *
 * Gated here, pure and headless (no live network, no Obsidian runtime):
 *   - THE ASK (#38, point 4): a vault with no `outlookAccounts` key, or a
 *     list that names only `default`, loads and saves data.json byte for
 *     byte, writes a note byte for byte as before, and lists the same
 *     secret keys;
 *   - an account id is validated and never transformed; `default` is
 *     reserved for today's sign-in;
 *   - the list: default first and always present, unusable records dropped;
 *   - the lookup: absent means default, an unlisted id is a blank disabled
 *     account and never the first mailbox;
 *   - the stamp: read as written, absent means default;
 *   - source scan: this step writes no stamp and lays no default under
 *     the key.
 *
 * The byte-for-byte tests also pass against the 0.16.2 bytes through
 * PLANNER_MAIN: they are the guard that this step changed nothing for a
 * one-account vault. The model tests go red there, because the functions
 * do not exist.
 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const T = require('./harness.cjs');

const PluginClass = require(T.__mainPath);
const { TFile } = T.__obsidian;

const code = () => fs.readFileSync(T.__mainPath, 'utf8');

// Invented throughout: a client id shape, a message id shape, two labels.
const CLIENT = '11111111-2222-3333-4444-555555555555';
const GRAPH_FEED = { id: 'outlook-graph', name: 'Outlook calendar', url: '', color: 2, enabled: true, kind: 'graph' };
const DEFAULT_ONLY = [{ accountId: 'default', label: 'Personal', enabled: true }];
const TWO = [{ accountId: 'default', label: 'Personal', enabled: true }, { accountId: 'work', label: 'Work', enabled: true }];

// What a 0.16.2 vault's data.json holds after one save: every default key
// in DEFAULT_SETTINGS order, the member's values over them, `_shadow` last
// (onload adds it after the defaults). `extra` lands after the known keys,
// which is where the plugin puts any key it did not lay down itself.
function storedText(extra) {
  const s = Object.assign({}, T.DEFAULT_SETTINGS, {
    outlookClientId: CLIENT,
    outlookScopes: 'Mail.Read Calendars.Read',
    calendars: [GRAPH_FEED],
    secretsInStore: true,
  }, extra || {}, { _shadow: { 'outlook:msg-1': { status: 'open', title: 'Renew the passport' } } });
  return JSON.stringify(s, null, 2);
}

// The plugin without Obsidian: the load path (adoptSettings, as onload runs
// it, in data-json mode with no store) and the save path (persistSettings),
// with saveData recorded as the JSON text that would reach disk.
async function loadAndSave(text) {
  const loaded = JSON.parse(text);
  const p = Object.create(PluginClass.prototype);
  p.secretStorage = null;
  p.secrets = p.vaultFor('secret-storage');
  assert.equal(p.secrets.mode, 'data-json');
  const adopted = T.adoptSettings(loaded, p.secrets);
  p.settings = adopted.settings;
  const saved = [];
  p.saveData = async (s) => { saved.push(JSON.stringify(s, null, 2)); };
  await p.persistSettings();
  assert.equal(saved.length, 1);
  return { changed: adopted.changed, saved: saved[0], settings: p.settings };
}

function noteWrittenWith(settings) {
  const calls = [];
  const p = Object.create(PluginClass.prototype);
  p.settings = settings;
  p.app = {
    vault: {
      getAbstractFileByPath: () => null,
      create: async (path, content) => { calls.push({ path, content }); },
    },
    fileManager: { processFrontMatter: async (file, fn) => { fn({}); } },
  };
  return p.createItemFile('02 Planner/Outlook', 'outlook', { id: 'msg-1', title: 'Renew the passport', priority: 4 }).then(() => calls[0]);
}

test('THE ASK (#38, point 4): no outlookAccounts key, or default only, loads and saves data.json byte for byte', async () => {
  assert.equal('outlookAccounts' in T.DEFAULT_SETTINGS, false, 'no default is laid under the key, so no save adds it');
  // No key at all: the vault every member has today.
  const without = storedText();
  assert.doesNotMatch(without, /outlookAccounts/);
  const a = await loadAndSave(without);
  assert.equal(a.changed, false, 'nothing to write back at load');
  assert.equal(a.saved, without, 'the save hands data.json back byte for byte');
  assert.equal('outlookAccounts' in a.settings, false, 'and the live settings carry no such key either');
  // A list that names only `default`: kept as typed, values untouched, no
  // record added, no record normalised into the file.
  const only = storedText({ outlookAccounts: DEFAULT_ONLY });
  const b = await loadAndSave(only);
  assert.equal(b.changed, false);
  assert.equal(b.saved, only);
  assert.deepEqual(b.settings.outlookAccounts, DEFAULT_ONLY);
});

test('THE ASK (#38, point 4): a note is written byte for byte as before, with no stamp, in both one-account shapes', async () => {
  const base = Object.assign({}, T.DEFAULT_SETTINGS, { outlookClientId: CLIENT, _shadow: {} });
  const without = await noteWrittenWith(base);
  const only = await noteWrittenWith(Object.assign({}, base, { outlookAccounts: DEFAULT_ONLY, _shadow: {} }));
  assert.equal(without.path, '02 Planner/Outlook/Renew the passport (outlook-msg-1).md');
  assert.equal(only.path, without.path);
  // synced_at is the clock at the write; the two writes are a tick apart.
  const clockless = (s) => s.replace(/^synced_at: .*$/m, 'synced_at: <now>');
  assert.match(without.content, /^synced_at: \d{4}-/m);
  assert.equal(clockless(only.content), clockless(without.content), 'the list that names only default changes nothing in the note');
  assert.doesNotMatch(without.content, /source_account/, 'no stamp on a default note');
  assert.match(without.content, /^source: outlook$/m);
  assert.match(without.content, /^external_id: "msg-1"$/m);
});

test('THE ASK (#38, point 4): the secret key list is the same with no key, with default only, and untouched by a second account', () => {
  const keys = (s) => T.secretSlots(s).map((slot) => [slot.id, slot.envKey, slot.ids]);
  const base = { calendars: [GRAPH_FEED] };
  assert.deepEqual(keys(Object.assign({}, base, { outlookAccounts: DEFAULT_ONLY })), keys(base));
  // A second account adds its own keys after the ones every member has
  // (part 2, the sign-in); the first five never move or change.
  assert.deepEqual(keys(Object.assign({}, base, { outlookAccounts: TWO })).slice(0, 5), keys(base).slice(0, 5), 'the five credentials by name are what they were');
  assert.deepEqual(Object.keys(T.SECRET_FIELDS), ['todoistToken', 'clickupToken', 'imapPassword', 'outlookRefreshToken', 'outlookAccessToken', 'outlookExpiresAt', 'outlookAccount']);
  assert.equal(T.settingsHoldSecrets({ outlookAccounts: TWO, calendars: [] }), false, 'the list holds no secret');
});

test('an account id is lowercase letters, digits and dashes: validated, never transformed', () => {
  assert.equal(T.OUTLOOK_DEFAULT_ACCOUNT, 'default');
  const accepted = ['default', 'work', 'work-2', 'a', '0', 'a-b-c', 'x'.repeat(32), 'a-' + 'x'.repeat(30)];
  for (const ok of accepted) assert.equal(T.outlookAccountId(ok), ok, ok);
  for (const bad of ['', ' ', 'Work', 'WORK', ' work', 'work ', 'wo rk', 'work_2', 'work.2', '-work', 'work-', 'a-', 'a--b', 'a@b', 'ünïcode', 'x'.repeat(33), 'a-' + 'x'.repeat(31), null, undefined, 7, {}, ['work']]) {
    assert.equal(T.outlookAccountId(bad), '', `rejected as given, never fixed up: ${JSON.stringify(bad)}`);
  }
  // The property that matters: an accepted id reaches a store key unchanged
  // (secretKey folds dash runs and strips a trailing dash, so a shape it
  // would fold is not an id), and no two accepted ids share an env key.
  const envKeys = new Set();
  for (const id of accepted) {
    const key = T.secretKey(`outlook-refresh-token-${id}`);
    assert.equal(key, `${T.SECRET_KEY_PREFIX}outlook-refresh-token-${id}`, `unchanged on the way into a key: ${id}`);
    assert.match(key, /^[a-z0-9-]+$/, 'the store alphabet, as the auth tests fake it');
    envKeys.add(T.envKeyFor(key));
  }
  assert.equal(envKeys.size, accepted.length, 'pairwise distinct env keys');
  assert.notEqual(T.secretKey('outlook-refresh-token-work-'), `${T.SECRET_KEY_PREFIX}outlook-refresh-token-work-`, 'the shape the rule refuses is the shape secretKey would fold');
  assert.equal(T.OUTLOOK_ACCOUNT_ID_RE.test('Work'), false, 'the rule is a regex the sign-in step reuses as is');
});

test('the list: default first and always present; unusable records fall out; enabled defaults to true; label falls back to the id', () => {
  const just = (s) => T.outlookAccountList(s);
  assert.deepEqual(just({}), [{ accountId: 'default', label: 'default', enabled: true }]);
  assert.deepEqual(just(null), just({}), 'no settings at all is one account');
  assert.deepEqual(just({ outlookAccounts: 'work' }), just({}), 'not a list: one account');
  assert.deepEqual(just({ outlookAccounts: [] }), just({}));
  assert.deepEqual(just({ outlookAccounts: DEFAULT_ONLY }), [{ accountId: 'default', label: 'Personal', enabled: true }], 'a default record adds a label and nothing else');
  assert.deepEqual(just({ outlookAccounts: TWO }), [
    { accountId: 'default', label: 'Personal', enabled: true },
    { accountId: 'work', label: 'Work', enabled: true },
  ]);
  // Default is first even when the member lists it second, and is added
  // when the member lists only the others.
  assert.deepEqual(just({ outlookAccounts: [{ accountId: 'work' }, { accountId: 'default', label: 'Personal' }] }).map((a) => a.accountId), ['default', 'work']);
  assert.deepEqual(just({ outlookAccounts: [{ accountId: 'work', label: 'Work' }] }), [
    { accountId: 'default', label: 'default', enabled: true },
    { accountId: 'work', label: 'Work', enabled: true },
  ]);
  // Unusable records: not an object, no id, a bad id, a repeated id.
  assert.deepEqual(just({ outlookAccounts: [null, 'work', 7, {}, { accountId: 'Work' }, { label: 'No id' }, { accountId: 'work' }, { accountId: 'work', label: 'Twice' }] }).map((a) => a.accountId), ['default', 'work']);
  assert.equal(just({ outlookAccounts: [{ accountId: 'work' }, { accountId: 'work', label: 'Twice' }] })[1].label, 'work', 'the first record with an id wins');
  // enabled: only an explicit false switches an account off; a label is trimmed and falls back to the id when blank.
  assert.equal(just({ outlookAccounts: [{ accountId: 'work', enabled: false }] })[1].enabled, false);
  assert.equal(just({ outlookAccounts: [{ accountId: 'work', enabled: 'no' }] })[1].enabled, true);
  assert.equal(just({ outlookAccounts: [{ accountId: 'default', enabled: false }] })[0].enabled, false, 'the default can be switched off too; what that means is the sync step\'s to decide');
  assert.equal(just({ outlookAccounts: [{ accountId: 'work', label: '  ' }] })[1].label, 'work');
  assert.equal(just({ outlookAccounts: [{ accountId: 'work', label: ' Work ' }] })[1].label, 'Work');
  // Pure: the list in the settings is never rewritten.
  const s = { outlookAccounts: [{ accountId: 'work', label: 'Work', folders: ['Inbox'] }, 'junk'] };
  const before = JSON.stringify(s);
  just(s);
  assert.equal(JSON.stringify(s), before, 'read, not normalised in place; room for later keys is left alone');
  assert.deepEqual(Object.keys(just(s)[1]), ['accountId', 'label', 'enabled'], 'and the read record carries only what this step defines');
});

test('the lookup: absent means default; an unlisted id is a blank disabled account, never the first mailbox', () => {
  const s = { outlookAccounts: [{ accountId: 'default', label: 'Personal' }, { accountId: 'work', label: 'Work' }, { accountId: 'old', label: 'Old', enabled: false }] };
  assert.deepEqual(T.outlookAccountById(s, undefined), { accountId: 'default', label: 'Personal', enabled: true });
  assert.deepEqual(T.outlookAccountById(s, null), T.outlookAccountById(s, undefined));
  assert.deepEqual(T.outlookAccountById(s, ''), T.outlookAccountById(s, undefined));
  assert.deepEqual(T.outlookAccountById(s, 'default'), T.outlookAccountById(s, undefined));
  assert.deepEqual(T.outlookAccountById(s, 'work'), { accountId: 'work', label: 'Work', enabled: true });
  assert.deepEqual(T.outlookAccountById(s, 'old'), { accountId: 'old', label: 'Old', enabled: false }, 'listed and switched off: still that account, still off');
  for (const stray of ['gone', 'Work', 'WORK', ' work', 'work ', 'default ']) {
    const a = T.outlookAccountById(s, stray);
    assert.equal(a.accountId, stray, 'the id as given, so the caller can name it');
    assert.equal(a.enabled, false, `not listed, so disabled: ${JSON.stringify(stray)}`);
    assert.notEqual(a.label, 'Personal', 'never the first mailbox');
  }
  assert.deepEqual(T.outlookAccountById({}, 'work'), { accountId: 'work', label: 'work', enabled: false }, 'no list at all: only default exists');
  assert.deepEqual(T.outlookAccountById(null, null), { accountId: 'default', label: 'default', enabled: true });
});

test('the stamp: source_account is read as written, and absent means default', () => {
  const fm = (extra) => Object.assign({ type: 'planner-item', source: 'outlook', external_id: 'msg-1', title: 'Renew the passport' }, extra || {});
  const plain = T.itemFromFrontmatter(fm(), 'p', 'b');
  assert.equal(plain.sourceAccount, null, 'no stamp: null on the item');
  assert.equal(T.itemAccountId(plain), 'default');
  assert.equal(T.itemFromFrontmatter(fm({ source_account: '' }), 'p', 'b').sourceAccount, null, 'an empty stamp is no stamp');
  assert.equal(T.itemFromFrontmatter(fm({ source_account: null }), 'p', 'b').sourceAccount, null);
  // Strings only: a list or a number the Properties editor typed is no stamp,
  // so `[work]` never matches the work account.
  assert.equal(T.itemFromFrontmatter(fm({ source_account: ['work'] }), 'p', 'b').sourceAccount, null);
  assert.equal(T.itemFromFrontmatter(fm({ source_account: 7 }), 'p', 'b').sourceAccount, null);
  const work = T.itemFromFrontmatter(fm({ source_account: 'work' }), 'p', 'b');
  assert.equal(work.sourceAccount, 'work');
  assert.equal(T.itemAccountId(work), 'work');
  assert.equal(work.id, 'msg-1', 'the rest of the item is what it was');
  assert.equal(work.source, 'outlook');
  // A hand-edited stamp is kept as written and resolves to a disabled
  // account, never to the first mailbox.
  const odd = T.itemFromFrontmatter(fm({ source_account: 'Work' }), 'p', 'b');
  assert.equal(odd.sourceAccount, 'Work');
  assert.equal(T.itemAccountId(odd), 'Work');
  assert.equal(T.outlookAccountById({ outlookAccounts: TWO }, T.itemAccountId(odd)).enabled, false);
  assert.equal(T.outlookAccountById({ outlookAccounts: TWO }, T.itemAccountId(work)).label, 'Work');
  // Not an Outlook thing by itself: every source reads the field the same way, and the helpers take any item shape.
  assert.equal(T.itemFromFrontmatter(fm({ source: 'todoist', source_account: 'x' }), 'p', 'b').sourceAccount, 'x');
  assert.equal(T.itemAccountId({}), 'default');
  assert.equal(T.itemAccountId(null), 'default');
  assert.equal(T.itemAccountId({ sourceAccount: '' }), 'default');
  assert.equal(T.itemAccountId({ sourceAccount: 'work' }), 'work');
});

test('source scan: the stamp has one reader and one writer (part 3, a further account\'s note only), and no default under outlookAccounts', () => {
  const c = code();
  const stampLines = c.split('\n').filter((l) => l.includes('source_account') && !/^\s*(\/\/|\*|\/\*)/.test(l));
  assert.equal(stampLines.length, 2, 'the reader in itemFromFrontmatter and the writer in createItemFile are the two lines of code that name the field');
  assert.match(stampLines[0], /^\s*sourceAccount: typeof fm\.source_account === 'string'/);
  assert.match(stampLines[1], /^\s*\.\.\.\(accountId && accountId !== OUTLOOK_DEFAULT_ACCOUNT \? \[`source_account: /, 'written only for a further account, in the same write as every other field');
  const defaults = c.slice(c.indexOf('const DEFAULT_SETTINGS = {'), c.indexOf('\n};', c.indexOf('const DEFAULT_SETTINGS = {')));
  assert.doesNotMatch(defaults, /outlookAccounts/, 'no default under the key, so no save adds it');
  // The list is read through one function and nowhere else, so the shape
  // lives in one place.
  const readers = c.split('\n').filter((l) => /\.outlookAccounts\b/.test(l) && !/^\s*(\/\/|\*|\/\*)/.test(l));
  assert.equal(readers.length, 1, 'outlookAccountList is the only reader of settings.outlookAccounts');
  assert.match(readers[0], /Array\.isArray\(s\.outlookAccounts\)/);
});
