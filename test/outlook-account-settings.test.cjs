/* More than one Microsoft account, part 4b of #38: the account list edited
 * from settings.
 *
 * Until here a second account meant a hand-edited `outlookAccounts` in
 * data.json. Now the Outlook section has an "Add account" row, and with two
 * or more accounts listed every account row carries its label and its
 * on/off switch, a further row a press-twice remove. The id is derived from
 * the label under part 1's rules and follows the label only while nothing
 * depends on it; removal signs the account out first and never touches a
 * note.
 *
 * Gated here, headless:
 *   - the pure rules: id from label, uniqueness, the label clash, the lock;
 *   - the plugin's four operations (add, rename, switch, remove) on a
 *     settings object, with save, model and calendar calls recorded;
 *   - the settings tab itself, rendered through a recording Setting stub
 *     (main.js re-required with the stub swapped in): a one-account vault
 *     sees its old row and one new Add row and nothing is saved; with two
 *     accounts the rows carry the controls and drive the operations;
 *   - source scan.
 *
 * Red first: every test but the first watched fail against the previous
 * part's bytes (PLANNER_MAIN), where none of the functions exists and the
 * tab has no Add row.
 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const T = require('./harness.cjs');

const PluginClass = require(T.__mainPath);
const code = () => fs.readFileSync(T.__mainPath, 'utf8');

const CLIENT = '11111111-2222-3333-4444-555555555555';
const GRAPH_FEED = { id: 'outlook-graph', name: 'Outlook calendar', url: '', color: 2, enabled: true, kind: 'graph' };
const WORK_FEED = { id: 'outlook-graph-work', name: 'Outlook calendar (Work)', url: '', color: 3, enabled: true, kind: 'graph', accountId: 'work' };
const TWO = [{ accountId: 'default', label: 'Personal' }, { accountId: 'work', label: 'Work', enabled: true, folders: ['Inbox'] }];
const item = (source, id, extra) => T.itemFromFrontmatter(Object.assign({ type: 'planner-item', source, external_id: id, title: id, status: 'open' }, extra || {}), `02 Planner/${source}/${id}.md`, id);

/* ---- 1. the id follows part 1's rules ---- */
test('THE GUARD and the id: from a label under the id rules, unique among the listed, default reserved', () => {
  const from = T.outlookAccountIdFromLabel;
  assert.equal(from('Work'), 'work');
  assert.equal(from('  My Work  Mail!  '), 'my-work-mail', 'runs of anything else become one dash');
  assert.equal(from('--Work--'), 'work', 'no leading or trailing dash');
  assert.equal(from('Ian_S.Personal'), 'ian-s-personal');
  assert.equal(from('Default'), '', 'the reserved id is never derived');
  assert.equal(from(''), '');
  assert.equal(from(null), '');
  assert.equal(from('!!!'), '', 'nothing usable: no id');
  const long = from('a'.repeat(30) + '-' + 'b'.repeat(10));
  assert.equal(long, 'a'.repeat(30) + '-b', 'cut at 32');
  assert.equal(from('a'.repeat(31) + '-' + 'b'.repeat(10)), 'a'.repeat(31), 'a dash left at the cut is removed');
  assert.ok(T.OUTLOOK_ACCOUNT_ID_RE.test(long));
  for (const l of ['Work', 'x y z', 'Ab-Cd', 'a'.repeat(40)]) assert.ok(T.outlookAccountId(from(l)), `${l}: passes the id rule`);
  const uniq = T.outlookUniqueAccountId;
  assert.equal(uniq('work', ['default']), 'work');
  assert.equal(uniq('work', ['default', 'work']), 'work-2');
  assert.equal(uniq('work', ['default', 'work', 'work-2']), 'work-3');
  assert.equal(uniq('', ['default']), 'account', 'an empty base starts from "account"');
  assert.equal(uniq('', ['default', 'account']), 'account-2');
  const full = 'a'.repeat(32);
  assert.equal(uniq(full, [full]), 'a'.repeat(30) + '-2', 'the suffix fits inside the cap');
  assert.ok(T.outlookAccountId(uniq(full, [full])));
  assert.equal(uniq('a'.repeat(31), ['a'.repeat(31)]), 'a'.repeat(30) + '-2', 'no doubled dash at the join');
});

/* ---- 2. a label another account holds ---- */
test('a label already held by another account is refused, compared trimmed and case-insensitively; the account\'s own label is not a clash', () => {
  const s = { outlookAccounts: TWO };
  assert.equal(T.outlookAccountLabelTaken(s, 'Work', 'default'), true);
  assert.equal(T.outlookAccountLabelTaken(s, '  work ', 'default'), true);
  assert.equal(T.outlookAccountLabelTaken(s, 'Work', 'work'), false, 'its own');
  assert.equal(T.outlookAccountLabelTaken(s, 'personal', 'work'), true, 'the default\'s label counts');
  assert.equal(T.outlookAccountLabelTaken(s, 'Home', 'work'), false);
  assert.equal(T.outlookAccountLabelTaken(s, '', 'work'), false);
  assert.equal(T.outlookAccountLabelTaken({}, 'default', 'x'), true, 'the default\'s label is its id when it has none');
});

/* ---- 3. the lock ---- */
test('the id is locked once anything depends on it: a sign-in (on the resolved copy), a shadow, a calendar feed or a note; the default always', () => {
  const inUse = T.outlookAccountIdInUse;
  const base = () => ({ outlookClientId: CLIENT, outlookAccounts: TWO, calendars: [GRAPH_FEED], _shadow: { 'outlook:d1': {}, 'outlook@old:o1': {} } });
  assert.equal(inUse(base(), 'default', []), true);
  assert.equal(inUse(base(), 'work', []), false, 'nothing depends on it yet');
  assert.equal(inUse(Object.assign(base(), { outlookRefreshToken__work: 'rt-w1' }), 'work', []), true, 'signed in');
  assert.equal(inUse(Object.assign(base(), { outlookScopes__work: 'Mail.Read' }), 'work', []), true, 'granted scopes remembered');
  assert.equal(inUse(Object.assign(base(), { outlookAccount__work: 'work-mailbox' }), 'work', []), true);
  assert.equal(inUse(Object.assign(base(), { outlookRefreshToken__work: '   ' }), 'work', []), false, 'blank is blank');
  assert.equal(inUse(Object.assign(base(), { _shadow: { 'outlook@work:w1': {} } }), 'work', []), true, 'a shadow under its key');
  assert.equal(inUse(Object.assign(base(), { _shadow: { 'outlook@workshop:w1': {} } }), 'work', []), false, 'another id with the same start is not it');
  assert.equal(inUse(Object.assign(base(), { calendars: [GRAPH_FEED, WORK_FEED] }), 'work', []), true, 'its calendar feed');
  assert.equal(inUse(base(), 'work', [item('outlook', 'w1', { source_account: 'work' })]), true, 'a note stamped with it');
  assert.equal(inUse(base(), 'work', [item('todoist', 't1', { source_account: 'work' })]), false, 'a stray stamp on a Todoist note is not Outlook\'s');
  assert.equal(inUse(base(), 'work', [item('outlook', 'd1'), null]), false);
  // The resolved copy is what the caller hands over: a token in the store
  // shows up there and nowhere on the live settings.
  const live = Object.assign(base(), { outlookRefreshToken__work: '' });
  const resolved = T.withSecrets(live, { available: () => true, get: () => 'rt-w1', has: () => true, list: () => [], mode: 'store' });
  assert.equal(inUse(live, 'work', []), false);
  assert.equal(inUse(resolved, 'work', []), true, 'read on the resolved copy, the keychain\'s token counts');
});

/* ---- the plugin under test ---- */
function plugin(settings, over) {
  const p = Object.create(PluginClass.prototype);
  const log = { saved: 0, model: 0, recomputed: 0, cleared: [], notices: [] };
  Object.assign(p, {
    settings, secrets: null, syncStatus: {}, syncStatusByAccount: {}, calendarDefsByFeed: {}, calendarDefs: null,
    app: { vault: { getAbstractFileByPath: () => { throw new Error('the vault is not read'); }, trash: () => { throw new Error('no note is touched'); }, delete: () => { throw new Error('no note is touched'); } }, fileManager: { renameFile: () => { throw new Error('no note is moved'); } } },
    saveSettings: async () => { log.saved += 1; },
    emitModelChanged: () => { log.model += 1; },
    recomputeCalendarDefs: () => { log.recomputed += 1; },
    outlookClearPending: (id) => { log.cleared.push(id); },
    withSecrets: () => Object.assign({}, p.settings),
    paths: () => { throw new Error('the plugin does not collect items itself'); },
  }, over || {});
  return { p, log };
}
async function withNotice(fn) {
  const Real = T.__obsidian.Notice;
  const notices = [];
  T.__obsidian.Notice = function (msg) { notices.push(String(msg)); };
  try { return await fn(notices); } finally { T.__obsidian.Notice = Real; }
}

/* ---- 4. add, rename, switch ---- */
test('THE ASK: Add lays the list down and pushes a labelled, switched-on record; the label may move the id until it is used; a duplicate or empty label is refused; the switch writes `enabled`', async () => {
  // A vault with no key: the first Add creates the list with the new
  // record alone (the default needs no record).
  const { p, log } = plugin({ outlookClientId: CLIENT, outlookRefreshToken: 'rt-d1', _shadow: {} });
  assert.equal(await p.addOutlookAccount(), 'account-2');
  assert.deepEqual(p.settings.outlookAccounts, [{ accountId: 'account-2', label: 'Account 2', enabled: true }]);
  assert.equal(log.saved, 1);
  assert.equal(log.model, 1, 'the tray and the cards re-read the list');
  assert.equal(await p.addOutlookAccount(), 'account-3');
  assert.deepEqual(T.outlookAccountList(p.settings).map((a) => [a.accountId, a.label, a.enabled]), [['default', 'default', true], ['account-2', 'Account 2', true], ['account-3', 'Account 3', true]]);
  // Renamed before use: the id follows the label, unique among the others,
  // and the old id's tray row is dropped.
  p.syncStatusByAccount = { 'account-2': { ok: true } };
  assert.deepEqual(await p.renameOutlookAccount('account-2', ' Work ', []), { ok: true, accountId: 'work' });
  assert.deepEqual(p.settings.outlookAccounts[0], { accountId: 'work', label: 'Work', enabled: true });
  assert.deepEqual(p.syncStatusByAccount, {});
  assert.deepEqual(await p.renameOutlookAccount('account-3', 'Work Mail', []), { ok: true, accountId: 'work-mail' });
  // The same label differently cased is the same label and is refused; a
  // different label whose id another listed account holds takes a suffix.
  assert.deepEqual(await p.renameOutlookAccount('work-mail', 'work', []), { ok: false, reason: 'duplicate' });
  assert.deepEqual(await p.renameOutlookAccount('work-mail', 'Work!', []), { ok: true, accountId: 'work-2' }, 'the label is new, the derived id is taken');
  assert.deepEqual(await p.renameOutlookAccount('work-2', 'Work Mail 2', []), { ok: true, accountId: 'work-mail-2' });
  // A label with nothing usable in it keeps the id and takes the label.
  assert.deepEqual(await p.renameOutlookAccount('work-mail-2', '***', []), { ok: true, accountId: 'work-mail-2' });
  assert.equal(p.settings.outlookAccounts[1].label, '***');
  // Locked by a note carrying the id: only the label moves.
  assert.deepEqual(await p.renameOutlookAccount('work', 'Client', [item('outlook', 'w1', { source_account: 'work' })]), { ok: true, accountId: 'work' });
  assert.equal(p.settings.outlookAccounts[0].label, 'Client');
  // Locked by a sign-in: the same, read on the resolved copy.
  p.settings.outlookRefreshToken__work = 'rt-w1';
  assert.deepEqual(await p.renameOutlookAccount('work', 'Client Two', []), { ok: true, accountId: 'work' });
  delete p.settings.outlookRefreshToken__work;
  // Refused: empty, duplicate (against the default's label too), unlisted.
  const saved = log.saved;
  assert.deepEqual(await p.renameOutlookAccount('work', '   ', []), { ok: false, reason: 'empty' });
  assert.deepEqual(await p.renameOutlookAccount('work', 'default', []), { ok: false, reason: 'duplicate' }, 'the default\'s label is its id here');
  assert.deepEqual(await p.renameOutlookAccount('gone', 'X', []), { ok: false, reason: 'unlisted' });
  assert.equal(log.saved, saved, 'a refusal saves nothing');
  // The default: a label creates its record, first in the list; its id never moves.
  assert.deepEqual(await p.renameOutlookAccount('default', 'Home', []), { ok: true, accountId: 'default' });
  assert.deepEqual(p.settings.outlookAccounts[0], { accountId: 'default', label: 'Home' });
  assert.deepEqual(await p.renameOutlookAccount('work', 'home', []), { ok: false, reason: 'duplicate' });
  // The switch.
  assert.equal(await p.setOutlookAccountEnabled('work', false), true);
  assert.equal(p.settings.outlookAccounts[1].enabled, false);
  assert.equal(T.outlookAccountById(p.settings, 'work').enabled, false);
  assert.equal(log.recomputed, 1, 'its calendar feed stops with it');
  assert.equal(await p.setOutlookAccountEnabled('work', true), true);
  assert.equal(p.settings.outlookAccounts[1].enabled, true);
  assert.equal(await p.setOutlookAccountEnabled('gone', false), false);
  // The default's switch creates its record when there is none.
  const { p: q, log: ql } = plugin({ outlookAccounts: [{ accountId: 'work', label: 'Work' }] });
  assert.equal(await q.setOutlookAccountEnabled('default', false), true);
  assert.deepEqual(q.settings.outlookAccounts, [{ accountId: 'default', enabled: false }, { accountId: 'work', label: 'Work' }]);
  assert.equal(ql.saved, 1);
  // Unknown keys on a record ride along untouched.
  const { p: r } = plugin({ outlookAccounts: [{ accountId: 'default', label: 'Personal' }, { accountId: 'work', label: 'Work', enabled: true, folders: ['Inbox'] }] });
  await r.renameOutlookAccount('work', 'Client', []);
  assert.deepEqual(r.settings.outlookAccounts[1], { accountId: 'client', label: 'Client', enabled: true, folders: ['Inbox'] });
});

/* ---- 5. remove ---- */
test('remove signs the account out first (its keys cleared, its suffixed fields gone), drops its feed, its pending sign-in and its tray row, and leaves every note alone; the default cannot be removed', async () => {
  const settings = {
    outlookClientId: CLIENT, outlookRefreshToken: 'rt-d1', outlookAccessToken: 'at-d1', outlookExpiresAt: '9', outlookAccount: 'me', outlookScopes: 'Mail.Read',
    outlookAccounts: [{ accountId: 'default', label: 'Personal' }, { accountId: 'work', label: 'Work', enabled: true, folders: ['Inbox'] }, { accountId: 'old', label: 'Old', enabled: false }],
    outlookRefreshToken__work: 'rt-w1', outlookAccessToken__work: 'at-w1', outlookExpiresAt__work: '9', outlookAccount__work: 'work-mailbox', outlookScopes__work: 'Mail.Read',
    outlookRefreshToken__old: 'rt-o1',
    calendars: [GRAPH_FEED, WORK_FEED, { id: 'cal-1', name: 'Family', url: 'https://calendar.example.org/basic.ics', color: 1, enabled: true, kind: 'ics' }],
    _shadow: { 'outlook:d1': {}, 'outlook@work:w1': {} },
  };
  const { p, log } = plugin(settings);
  p.syncStatus = { outlook: { ok: false, message: 'Work: down' } };
  p.syncStatusByAccount = { default: { ok: true }, work: { ok: false } };
  p.calendarDefsByFeed = { 'outlook-graph': [], 'outlook-graph-work': [] };
  await withNotice(async (notices) => {
    assert.equal(await p.removeOutlookAccount('default'), false, 'the default cannot be removed');
    assert.equal(await p.removeOutlookAccount('gone'), false);
    assert.equal(await p.removeOutlookAccount('Not An Id'), false);
    assert.equal(log.saved, 0);
    assert.equal(await p.removeOutlookAccount('work'), true);
    assert.deepEqual(p.settings.outlookAccounts, [{ accountId: 'default', label: 'Personal' }, { accountId: 'old', label: 'Old', enabled: false }], 'the record is gone, the others as they were');
    for (const f of ['outlookRefreshToken', 'outlookAccessToken', 'outlookExpiresAt', 'outlookAccount', 'outlookScopes']) {
      assert.equal(`${f}__work` in p.settings, false, `${f}__work left data.json`);
      assert.equal(p.settings[f], settings[f] === undefined ? undefined : p.settings[f], 'the default\'s flat field is not the one cleared');
    }
    assert.equal(p.settings.outlookRefreshToken, 'rt-d1', 'the default\'s token untouched');
    assert.equal(p.settings.outlookRefreshToken__old, 'rt-o1', 'the other account\'s token untouched');
    assert.deepEqual(log.cleared, ['work'], 'its pending sign-in is forgotten');
    assert.deepEqual(p.syncStatusByAccount, { default: { ok: true } }, 'its tray row is gone');
    assert.equal(p.syncStatus.outlook, undefined, 'the folded row that may have named it is dropped, as a sign-out does');
    assert.deepEqual(p.settings.calendars.map((f) => f.id), ['outlook-graph', 'cal-1'], 'its feed goes; the default\'s and the iCal feed stay');
    assert.deepEqual(p.settings.calendars[0], GRAPH_FEED, 'the default\'s feed byte for byte');
    assert.equal(log.recomputed, 1, 'its events leave the board now');
    assert.equal(log.saved, 1);
    assert.deepEqual(p.settings._shadow, { 'outlook:d1': {}, 'outlook@work:w1': {} }, 'shadows are not touched; the next sync of the list prunes nothing for an account that has no run');
    assert.ok(Array.isArray(notices));
  });
  // The Notice (main.js binds the class at load, so its text is read here).
  const removal = code().slice(code().indexOf('async removeOutlookAccount('), code().indexOf('  markSyncWrite(fileOrPath) {'));
  assert.match(removal, /new Notice\(`Planner: removed the Outlook account "\$\{label\}" and its sign-in\. Its notes stay in the vault and show under "Not in the account list" in the tray\.`, 8000\);/);
  // No vault method was called (the fixture throws on any): the notes stay
  // where they are and show under the tray's unlisted section.
  const items = [item('outlook', 'w1', { source_account: 'work' }), item('outlook', 'd1')];
  const parts = T.traySourceSections(p.settings, 'outlook', items, {});
  assert.deepEqual(parts.map((x) => [x.key, x.unlisted]), [['outlook', false], ['outlook@old', false], ['outlook@work', true]]);
  // In store mode the same clearing goes through the store: the four keys of
  // that account are blanked (sign-out writes the empty string, as
  // clearOutlookTokens always has) and nobody else's is touched.
  const sets = [];
  const vault = { mode: 'store', available: () => true, has: () => true, get: () => 'x', list: () => [], set: (k, v) => { sets.push([k, v]); return true; } };
  const { p: q } = plugin({ outlookClientId: CLIENT, outlookAccounts: [{ accountId: 'default' }, { accountId: 'work', label: 'Work' }], calendars: [], _shadow: {}, secretsInStore: true }, { secrets: vault });
  await withNotice(async () => { assert.equal(await q.removeOutlookAccount('work'), true); });
  assert.deepEqual(sets.map(([k]) => k).sort(), ['icor-for-life-planner-outlook-access-token-work', 'icor-for-life-planner-outlook-account-work', 'icor-for-life-planner-outlook-expires-at-work', 'icor-for-life-planner-outlook-refresh-token-work'], 'the four keys of that account, none of the default\'s');
  assert.ok(sets.every(([, v]) => v === ''), 'blanked');
});

/* ---- the settings tab, rendered ---- */
class FakeEl {
  constructor(tag) {
    this.tagName = String(tag).toLowerCase(); this.children = []; this.attrs = {}; this._classes = []; this._text = ''; this.listeners = {}; this.dataset = {}; this.value = ''; this.disabled = false;
    const self = this;
    this.classList = {
      add: (...c) => { for (const x of c) if (x && !self._classes.includes(x)) self._classes.push(x); },
      remove: (...c) => { self._classes = self._classes.filter((x) => !c.includes(x)); },
      contains: (c) => self._classes.includes(c),
      toggle: (c, f) => { const on = f === undefined ? !self._classes.includes(c) : !!f; if (on) self.classList.add(c); else self.classList.remove(c); return on; },
    };
  }
  get className() { return this._classes.join(' '); }
  set className(v) { this._classes = String(v || '').split(/\s+/).filter(Boolean); }
  get textContent() { return this._text + this.children.map((c) => (c instanceof FakeEl ? c.textContent : c.text)).join(''); }
  set textContent(v) { this._text = String(v == null ? '' : v); this.children = []; }
  setAttribute(k, v) { this.attrs[k] = String(v); }
  getAttribute(k) { return Object.prototype.hasOwnProperty.call(this.attrs, k) ? this.attrs[k] : null; }
  removeAttribute(k) { delete this.attrs[k]; }
  appendChild(c) { this.children.push(c); return c; }
  append(...c) { for (const x of c) this.children.push(typeof x === 'string' ? { text: x } : x); }
  appendText(t) { this.children.push({ text: String(t) }); }
  prepend(c) { this.children.unshift(c); }
  addEventListener(t, fn) { (this.listeners[t] = this.listeners[t] || []).push(fn); }
  removeEventListener() { }
  querySelector() { return null; }
  querySelectorAll() { return []; }
  focus() { } select() { } remove() { }
  empty() { this.children = []; this._text = ''; }
  addClass(...c) { this.classList.add(...c); }
  removeClass(...c) { this.classList.remove(...c); }
  toggleClass(c, on) { this.classList.toggle(c, on); }
  hasClass(c) { return this.classList.contains(c); }
  setText(t) { this.textContent = t; }
  createEl(tag, o) {
    const e = new FakeEl(tag);
    const opt = typeof o === 'string' ? { cls: o } : (o || {});
    if (opt.cls) e.classList.add(...(Array.isArray(opt.cls) ? opt.cls : String(opt.cls).split(/\s+/)));
    if (opt.text != null) e.textContent = opt.text;
    if (opt.attr) for (const [k, v] of Object.entries(opt.attr)) e.setAttribute(k, v);
    if (opt.href != null) e.setAttribute('href', opt.href);
    if (opt.type != null) e.setAttribute('type', opt.type);
    if (opt.placeholder != null) e.setAttribute('placeholder', opt.placeholder);
    if (opt.value != null) e.value = opt.value;
    this.appendChild(e);
    return e;
  }
  createDiv(o) { return this.createEl('div', o); }
  createSpan(o) { return this.createEl('span', o); }
}
// Obsidian's Setting, recording: every component is an element under
// controlEl with `data-kind`, its handlers kept so a test can press it.
class Component {
  constructor(kind, host) {
    this.kind = kind; this.handlers = {};
    this.el = host.controlEl.createEl(kind === 'button' ? 'button' : kind === 'toggle' ? 'div' : kind === 'text' ? 'input' : 'div', { attr: { 'data-kind': kind } });
    this.inputEl = this.el; this.toggleEl = this.el; this.buttonEl = this.el; this.extraSettingsEl = this.el; this.selectEl = this.el; this.sliderEl = this.el;
  }
  setValue(v) { this.el.value = v; this.el.setAttribute('data-value', String(v)); return this; }
  getValue() { return this.el.value; }
  setPlaceholder(v) { this.el.setAttribute('placeholder', v); return this; }
  setDisabled(v) { this.el.disabled = !!v; this.el.setAttribute('data-disabled', v ? '1' : '0'); return this; }
  setButtonText(v) { this.el.textContent = v; return this; }
  setCta() { this.el.addClass('mod-cta'); return this; }
  removeCta() { this.el.removeClass('mod-cta'); return this; }
  setWarning() { return this; }
  setIcon(v) { this.el.setAttribute('data-icon', v); return this; }
  setTooltip(v) { this.el.setAttribute('title', v); return this; }
  setDynamicTooltip() { return this; } setLimits() { return this; } setInstant() { return this; }
  addOption(k, v) { this.el.createEl('option', { text: v, attr: { value: k } }); return this; }
  addOptions(o) { for (const [k, v] of Object.entries(o)) this.addOption(k, v); return this; }
  onChange(fn) { this.handlers.change = fn; return this; }
  onClick(fn) { this.handlers.click = fn; return this; }
  then(fn) { fn(this); return this; }
}
class FakeSetting {
  constructor(containerEl) {
    this.settingEl = containerEl.createDiv({ cls: 'setting-item' });
    this.infoEl = this.settingEl.createDiv({ cls: 'setting-item-info' });
    this.nameEl = this.infoEl.createDiv({ cls: 'setting-item-name' });
    this.descEl = this.infoEl.createDiv({ cls: 'setting-item-description' });
    this.controlEl = this.settingEl.createDiv({ cls: 'setting-item-control' });
    this.components = [];
    FakeSetting.all.push(this);
  }
  setName(v) { this.nameEl.textContent = v; return this; }
  setDesc(v) { if (typeof v === 'string') this.descEl.textContent = v; else { this.descEl.empty(); this.descEl.appendChild(v); } return this; }
  setClass(c) { this.settingEl.addClass(c); return this; }
  setHeading() { this.settingEl.addClass('setting-item-heading'); return this; }
  setTooltip() { return this; } setDisabled() { return this; }
  _add(kind, cb) { const c = new Component(kind, this); this.components.push(c); cb(c); return this; }
  addText(cb) { return this._add('text', cb); } addTextArea(cb) { return this._add('textarea', cb); }
  addToggle(cb) { return this._add('toggle', cb); } addButton(cb) { return this._add('button', cb); }
  addExtraButton(cb) { return this._add('extra', cb); } addDropdown(cb) { return this._add('dropdown', cb); }
  addSlider(cb) { return this._add('slider', cb); } addMomentFormat(cb) { return this._add('moment', cb); } addSearch(cb) { return this._add('search', cb); }
  then(fn) { fn(this); return this; }
}
FakeSetting.all = [];
// main.js again, with the recording Setting in the stub: the harness hands
// the same stub object to every require of 'obsidian', and main.js
// destructures it at load, so a fresh load sees the swap.
function loadWithFakeSetting() {
  const stub = T.__obsidian;
  const prev = stub.Setting;
  stub.Setting = FakeSetting;
  const key = require.resolve(T.__mainPath);
  const cached = require.cache[key];
  delete require.cache[key];
  try { return require(T.__mainPath).__test; } finally { stub.Setting = prev; if (cached) require.cache[key] = cached; else delete require.cache[key]; }
}
const TAB = loadWithFakeSetting();
function renderTab(settings, over) {
  const calls = { renamed: [], switched: [], removed: [], added: 0, displayed: 0, saved: 0 };
  const hadDoc = 'document' in globalThis; const hadWin = 'window' in globalThis;
  const prevDoc = globalThis.document; const prevWin = globalThis.window;
  globalThis.document = { createElement: (t) => new FakeEl(t), createElementNS: (ns, t) => new FakeEl(t), createTextNode: (t) => ({ text: String(t) }), body: new FakeEl('body'), querySelectorAll: () => [], activeElement: null };
  globalThis.window = { setTimeout: () => 7, clearTimeout: () => { }, open: () => { } };
  const plugin = Object.assign({
    settings, secrets: { mode: 'data-json', available: () => false, has: () => false },
    manifest: { id: 'icor-for-life-planner', version: '0.0.0' },
    app: { vault: { getAbstractFileByPath: () => null, getRoot: () => ({ children: [] }), adapter: {} }, workspace: {} },
    withSecrets: () => Object.assign({}, settings, { _live: settings }),
    syncStatus: {}, syncStatusByAccount: {}, calendarStatus: null, calendarDefsByFeed: {}, lastSyncAt: null,
    paths: () => ({ root: '02 Planner' }),
    saveSettings: async () => { calls.saved += 1; },
    emitModelChanged() { }, recomputeCalendarDefs() { }, envStore: { path: '.env', load: async () => { } },
    habitsFolder: () => '04 My Life/Habits', secretHolders: () => [], importCandidates: () => [], routines: [], habits: [], secretStorage: null,
    setupNextBadge() { }, removeNextBadge() { }, openImportHabits() { }, setRoutineActive() { }, setSecretsBackend() { }, setEnvFilePath() { }, moveSecretsFrom() { }, ensureFolders: async () => { }, syncNow() { },
    outlookSignIn() { }, outlookSignOut: async () => { },
    addOutlookAccount: async () => { calls.added += 1; return 'account-2'; },
    renameOutlookAccount: async (id, label, items) => { calls.renamed.push([id, label, items.length]); return label.trim().toLowerCase() === 'personal' ? { ok: false, reason: 'duplicate' } : label.trim() ? { ok: true, accountId: id } : { ok: false, reason: 'empty' }; },
    setOutlookAccountEnabled: async (id, on) => { calls.switched.push([id, on]); return true; },
    removeOutlookAccount: async (id) => { calls.removed.push(id); return true; },
  }, over || {});
  const tab = Object.create(TAB.IcorPlannerSettingTab.prototype);
  tab.plugin = plugin; tab.app = plugin.app;
  tab.containerEl = new FakeEl('div');
  tab.display = function () { calls.displayed += 1; };
  FakeSetting.all = [];
  try { TAB.IcorPlannerSettingTab.prototype.display.call(tab); } finally {
    if (hadDoc) globalThis.document = prevDoc; else delete globalThis.document;
    if (hadWin) globalThis.window = prevWin; else delete globalThis.window;
  }
  const rows = FakeSetting.all;
  const named = (prefix) => rows.filter((r) => r.nameEl.textContent.startsWith(prefix));
  return { rows, named, calls, plugin };
}
const kinds = (row) => row.components.map((c) => c.kind);
// From the defaults, as a loaded vault is: the tab normalises a few keys in
// place on render (routineDefaults, the calendar list), which is its
// existing behaviour and not this part's.
const ONE = () => Object.assign({}, T.DEFAULT_SETTINGS, { outlookClientId: CLIENT, outlookRefreshToken: 'rt-d1', calendars: [], _shadow: {} });

/* ---- 6. one account: the old row, one Add row, nothing saved ---- */
test('a one-account vault sees its Microsoft account row as it was and one new "Add account" row; the render saves nothing and lays no key down', () => {
  for (const s of [ONE(), Object.assign(ONE(), { outlookAccounts: [] }), Object.assign(ONE(), { outlookAccounts: [{ accountId: 'default', label: 'Personal' }] })]) {
    const before = JSON.stringify(s);
    const { named, calls } = renderTab(s);
    const acct = named('Microsoft account');
    assert.equal(acct.length, 1, 'one account row');
    assert.equal(acct[0].nameEl.textContent, 'Microsoft account', 'no label in the name with one account');
    assert.deepEqual(kinds(acct[0]), ['button', 'button'], 'sign in and sign out, no name field, no switch, no remove');
    const add = named('Add account');
    assert.equal(add.length, 1);
    assert.deepEqual(kinds(add[0]), ['button']);
    assert.equal(add[0].components[0].el.textContent, 'Add');
    assert.equal(add[0].components[0].el.getAttribute('data-add-account'), '1');
    assert.match(add[0].descEl.textContent, /^A second Microsoft account gets its own row/);
    assert.equal(calls.saved, 0, 'rendering saves nothing');
    assert.equal(JSON.stringify(s), before, 'the settings object is what it was: no key laid down');
    // Pressing Add is the one way in.
    add[0].components[0].handlers.click();
    assert.equal(calls.added, 1);
  }
});

/* ---- 7. two accounts: the rows carry the controls and drive the operations ---- */
test('with two accounts every row carries its name and its switch, a further row its press-twice remove, and the controls call the plugin\'s operations', async () => {
  // The handlers pressed below run after the render and reach for window
  // (the remove's disarm timer), as they would in the app.
  const hadWindow = 'window' in globalThis;
  if (!hadWindow) globalThis.window = { setTimeout: () => 7, clearTimeout: () => { }, open: () => { } };
  try { await twoAccountRows(); } finally { if (!hadWindow) delete globalThis.window; }
});
async function twoAccountRows() {
  const s = Object.assign(ONE(), { outlookAccounts: TWO, outlookRefreshToken__work: 'rt-w1' });
  const { named, calls } = renderTab(s);
  const rows = named('Microsoft account: ');
  assert.deepEqual(rows.map((r) => r.nameEl.textContent), ['Microsoft account: Personal', 'Microsoft account: Work']);
  assert.deepEqual(kinds(rows[0]), ['text', 'toggle', 'button', 'button'], 'the default: name, switch, sign in, sign out; no remove');
  assert.deepEqual(kinds(rows[1]), ['text', 'toggle', 'button', 'button', 'extra']);
  const [dName, dToggle] = rows[0].components;
  const [wName, wToggle, wIn, wOut, wRemove] = rows[1].components;
  assert.equal(dName.el.value, 'Personal');
  assert.equal(dName.el.getAttribute('data-account-label'), 'default');
  assert.equal(wName.el.value, 'Work');
  assert.equal(wName.el.getAttribute('data-account-label'), 'work');
  assert.equal(wName.el.getAttribute('aria-label'), 'Name of the Microsoft account Work');
  assert.equal(wToggle.el.getAttribute('data-value'), 'true');
  assert.equal(wToggle.el.getAttribute('aria-label'), 'Sync the Microsoft account Work');
  assert.equal(wIn.el.textContent, 'Sign in again', 'signed in');
  assert.equal(wOut.el.getAttribute('data-disabled'), '0');
  assert.equal(wRemove.el.getAttribute('data-icon'), 'trash');
  assert.equal(wRemove.el.getAttribute('aria-label'), 'Remove the Microsoft account Work');
  assert.match(named('Add account')[0].descEl.textContent, /Removing a row signs it out; its notes stay\./);
  // Rename: the vault's items ride along for the lock; a duplicate is
  // refused and said in the row; a good name renames the row at once.
  await wName.handlers.change('Personal');
  assert.deepEqual(calls.renamed, [['work', 'Personal', 0]]);
  assert.equal(rows[1].descEl.textContent, 'Another account already has this name.');
  await wName.handlers.change('Client');
  assert.equal(rows[1].nameEl.textContent, 'Microsoft account: Client');
  await wName.handlers.change('  ');
  assert.equal(rows[1].descEl.textContent, 'A name is needed.');
  await dName.handlers.change('Home');
  assert.deepEqual(calls.renamed.map((r) => r[0]), ['work', 'work', 'work', 'default']);
  // The switch.
  await wToggle.handlers.change(false);
  await dToggle.handlers.change(false);
  assert.deepEqual(calls.switched, [['work', false], ['default', false]]);
  // Remove: armed on the first press, done on the second, then a re-render.
  await wRemove.handlers.click();
  assert.deepEqual(calls.removed, []);
  assert.equal(wRemove.el.getAttribute('data-icon'), 'alert-triangle');
  assert.equal(wRemove.el.getAttribute('aria-label'), 'Press again to sign out and remove');
  assert.equal(calls.displayed, 0);
  await wRemove.handlers.click();
  assert.deepEqual(calls.removed, ['work']);
  assert.equal(calls.displayed, 1);
  assert.equal(calls.saved, 0, 'the tab itself saves nothing; the operations do');
  // Add re-renders and hands focus to the new row's name.
  await named('Add account')[0].components[0].handlers.click();
  assert.equal(calls.added, 1);
  assert.equal(calls.displayed, 2);
  // The items for the lock are collected once per render and only with more
  // than one account: pinned in the source scan below (the tab reads the
  // vault for other rows too, so a call count would not isolate it).
}

/* ---- 8. source scan ---- */
test('source scan: the list has one door, the removal clears the keys the way sign-out does, the Add row follows the account rows, and no text sends the member to data.json', () => {
  const c = code();
  const readers = c.split('\n').filter((l) => /\.outlookAccounts\b/.test(l) && !/^\s*(\/\/|\*|\/\*)/.test(l));
  assert.equal(readers.length, 1, 'outlookAccountRawList is the one line that touches the key');
  assert.match(readers[0], /return Array\.isArray\(s\.outlookAccounts\) \? s\.outlookAccounts : \(create \? \(s\.outlookAccounts = \[\]\) : \[\]\);/);
  assert.equal((c.match(/outlookAccountRawList\(this\.settings, true\)/g) || []).length, 2, 'the record lookup and Add lay the list down; nothing else does');
  assert.equal((c.match(/clearOutlookTokens\(\{ live: this\.settings, vault: this\.secrets, account: id \}\);/g) || []).length, 2, 'sign-out and removal clear an account\'s keys the same way');
  assert.match(c, /clearOutlookTokens\(\{ live: this\.settings, vault: this\.secrets, account: id \}\);\n\s*for \(const f of OUTLOOK_ACCOUNT_FIELDS\) delete this\.settings\[outlookAccountField\(id, f\)\];\n\s*this\.outlookClearPending\(id\);\n\s*delete this\.syncStatusByAccount\[id\];\n\s*delete this\.syncStatus\.outlook;\n\s*this\.settings\.calendars = calendarFeeds\(this\.settings\)\.filter/, 'removal: keys, fields, pending, rows, feed, in that order, before the record goes');
  assert.match(c, /if \(!id \|\| id === OUTLOOK_DEFAULT_ACCOUNT\) return false;/, 'the default is refused first');
  const removal = c.slice(c.indexOf('async removeOutlookAccount('), c.indexOf('  markSyncWrite(fileOrPath) {'));
  assert.ok(removal.length > 200 && removal.includes('list.splice(at, 1);'), 'the removal body');
  assert.doesNotMatch(removal, /vault\.(trash|delete|modify|rename)|processFrontMatter|renameFile/, 'removal touches no note');
  const tab = c.slice(c.indexOf('class IcorPlannerSettingTab'));
  assert.ok(tab.indexOf('for (const account of accounts.slice(1))') < tab.indexOf(".setName('Add account')"), 'Add after the rows');
  assert.ok(tab.indexOf(".setName('Add account')") < tab.indexOf("setName('Manage or revoke access')"), 'and before the revoke row');
  assert.match(tab, /const noteItems = accounts\.length > 1 \? collectItems\(this\.plugin\.app, this\.plugin\.paths\(\)\.root\) : \[\];/);
  assert.match(tab, /if \(accounts\.length > 1\) accountControls\(acct, accounts\[0\]\);/, 'the default row gains its controls only with more than one account');
  assert.match(tab, /setTooltip\('Press again to sign out and remove'\)/);
  assert.doesNotMatch(c, /data\.json first/, 'no Notice sends the member to data.json');
  assert.match(c, /add it under Outlook in the plugin's settings first\./);
  const buf = fs.readFileSync(T.__mainPath);
  let cr = 0;
  for (const b of buf) if (b === 13) cr += 1;
  assert.equal(cr, 0);
  assert.equal(path.basename(T.__mainPath), 'main.js');
});
