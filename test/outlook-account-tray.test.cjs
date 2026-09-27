/* More than one Microsoft account, part 4 of #38: the tray by account.
 *
 * With two or more accounts listed in `outlookAccounts`, the tray's
 * "unscheduled by source" loop renders one OUTLOOK section per account, in
 * list order (default first), headed by the account's label (its id when it
 * has none), each with its own count, its own collapse state and its OWN
 * run's status, so a mailbox that is not signed in or unreachable says so
 * under its own head. With zero or one account it renders exactly what it
 * always has, and the per-account status map stays empty.
 *
 * Two kinds of gate. The pure ones read traySourceSections, the decision the
 * renderer takes. The rendered ones drive renderSync itself through a small
 * DOM stand-in (Obsidian's createDiv / createEl / createSpan / addClass on
 * top of the plain DOM the cards use) and serialise the tree, so the head
 * text, the counts, the card membership, the notes and the collapse class
 * are asserted on what the member would see, not on a regex alone.
 *
 * Red first. Every test but the first was watched fail against the previous
 * part's bytes (PLANNER_MAIN at the calendar part): traySourceSections does
 * not exist there and the tray renders one section. Test 1 is the no-change
 * gate and is green there by design.
 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const T = require('./harness.cjs');

const PluginClass = require(T.__mainPath);
const code = () => fs.readFileSync(T.__mainPath, 'utf8');

/* ---- a DOM small enough to read ---- */
class FakeEl {
  constructor(tag) {
    this.tagName = String(tag).toLowerCase();
    this.children = [];
    this.attrs = {};
    this._classes = [];
    this._text = '';
    this.listeners = {};
    this.dataset = {};
    this.hidden = false;
    this.disabled = false;
    this.draggable = false;
    this.value = '';
    const self = this;
    this.classList = {
      add: (...c) => { for (const x of c) if (x && !self._classes.includes(x)) self._classes.push(x); },
      remove: (...c) => { self._classes = self._classes.filter((x) => !c.includes(x)); },
      contains: (c) => self._classes.includes(c),
      toggle: (c, force) => {
        const on = force === undefined ? !self._classes.includes(c) : !!force;
        if (on) self.classList.add(c); else self.classList.remove(c);
        return on;
      },
    };
  }
  get className() { return this._classes.join(' '); }
  set className(v) { this._classes = String(v || '').split(/\s+/).filter(Boolean); }
  get textContent() {
    return this._text + this.children.map((c) => (c instanceof FakeEl ? c.textContent : c.text)).join('');
  }
  set textContent(v) { this._text = String(v == null ? '' : v); this.children = []; }
  setAttribute(k, v) { this.attrs[k] = String(v); }
  getAttribute(k) { return Object.prototype.hasOwnProperty.call(this.attrs, k) ? this.attrs[k] : null; }
  hasAttribute(k) { return Object.prototype.hasOwnProperty.call(this.attrs, k); }
  appendChild(c) { this.children.push(c); return c; }
  appendText(t) { this.children.push({ text: String(t) }); }
  addEventListener(type, fn) { (this.listeners[type] = this.listeners[type] || []).push(fn); }
  removeEventListener() { }
  remove() { }
  focus() { }
  setSelectionRange() { }
  /* Obsidian's helpers, the subset the tray uses. */
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
  addClass(...c) { this.classList.add(...c); }
  removeClass(...c) { this.classList.remove(...c); }
  toggleClass(c, on) { this.classList.toggle(c, on); }
  hasClass(c) { return this.classList.contains(c); }
  empty() { this.children = []; this._text = ''; }
  /* A readable projection: what a diff would show. Listeners by type only. */
  toJSON() {
    const out = { tag: this.tagName };
    if (this._classes.length) out.cls = this._classes.join(' ');
    if (Object.keys(this.attrs).length) out.attrs = this.attrs;
    if (this._text) out.text = this._text;
    const ev = Object.keys(this.listeners).sort().map((k) => `${k}:${this.listeners[k].length}`);
    if (ev.length) out.on = ev;
    if (this.draggable) out.draggable = true;
    if (this.children.length) out.kids = this.children.map((c) => (c instanceof FakeEl ? c.toJSON() : { text: c.text }));
    return out;
  }
}

function withDom(fn) {
  const hadDoc = Object.prototype.hasOwnProperty.call(global, 'document');
  const hadWin = Object.prototype.hasOwnProperty.call(global, 'window');
  const prevDoc = global.document;
  const prevWin = global.window;
  global.document = {
    createElement: (t) => new FakeEl(t),
    createElementNS: (ns, t) => new FakeEl(t),
    createTextNode: (t) => ({ text: String(t) }),
    body: new FakeEl('body'),
    querySelectorAll: () => [],
    activeElement: null,
  };
  global.window = { setTimeout: () => 0, clearTimeout: () => { }, open: () => { } };
  try { return fn(); } finally {
    if (hadDoc) global.document = prevDoc; else delete global.document;
    if (hadWin) global.window = prevWin; else delete global.window;
  }
}

/* ---- fixtures, invented throughout: a client id shape, token shapes, two labels ---- */
const TODAY = '2026-09-27';
const CLIENT = '11111111-2222-3333-4444-555555555555';
const OUTLOOK = '02 Planner/Outlook';
function item(source, id, extra) {
  const fm = Object.assign({ type: 'planner-item', source, external_id: id, title: `${source} ${id}`, status: 'open' }, extra || {});
  return T.itemFromFrontmatter(fm, `02 Planner/${source}/${id}.md`, String(id));
}
// Twelve notes with no source_account, eight stamped `work`: two mailboxes
// in one vault. One Todoist item so a second source is in the loop too.
function fixtureItems() {
  const out = [];
  for (let i = 1; i <= 12; i += 1) out.push(item('outlook', `d-${i}`, { priority: (i % 4) + 1 }));
  for (let i = 1; i <= 8; i += 1) out.push(item('outlook', `w-${i}`, { source_account: 'work', priority: (i % 3) + 1 }));
  out.push(item('todoist', 't-1'));
  return out;
}
const SIGNED_IN = { outlookClientId: CLIENT, outlookRefreshToken: 'rt-d1', outlookTenant: 'consumers', todoistToken: 'tok' };
const ONE_ACCOUNT = Object.assign({}, SIGNED_IN);
const TWO_ACCOUNTS = Object.assign({}, SIGNED_IN, {
  outlookAccounts: [
    { accountId: 'default', label: 'Personal' },
    { accountId: 'work', label: 'Work' },
  ],
  outlookRefreshToken__work: 'rt-w1',
});
const with2 = (over, accounts) => Object.assign({}, TWO_ACCOUNTS, over || {}, accounts ? { outlookAccounts: accounts } : {});

function trayFor(settings, items, opts) {
  const o = opts || {};
  const plugin = {
    withSecrets: () => Object.assign({}, settings),
    syncStatus: o.syncStatus || {},
    syncStatusByAccount: o.byAccount || {},
    settings: { subtaskChecklist: false },
    app: { vault: { getAbstractFileByPath: () => null }, workspace: { getLeaf: () => ({ openFile() { } }) } },
    openPluginSettings() { },
    addManualItem: async () => null,
    toggleDoneLocal() { },
  };
  const view = Object.create(T.PlannerTrayView.prototype);
  view.plugin = plugin;
  view.collapsed = o.collapsed || {};
  view.composerOpen = false;
  view.composerDraft = '';
  view._composerWantsFocus = false;
  view.index = null;
  view.expanded = new Set();
  const root = withDom(() => {
    const el = new FakeEl('div');
    view.renderSync(el, items, TODAY);
    return el;
  });
  return { root, view };
}
// The source sections in render order: [{ head, mark, count, collapsed, cards, note, connect }].
function sections(root) {
  return root.children
    .filter((c) => c instanceof FakeEl && c.hasClass('iplan-tray-section') && c.children[0] && c.children[0].hasClass('is-clickable'))
    .map((sec) => {
      const head = sec.children[0];
      const body = sec.children[1];
      const note = body.children.find((c) => c instanceof FakeEl && c.hasClass('iplan-tray-note'));
      return {
        head: head.children.filter((c) => c instanceof FakeEl && c.tagName === 'span' && !c.hasClass('iplan-source-mark') && !c.hasClass('iplan-tray-count')).map((c) => c.textContent).join(''),
        mark: head.children[0].className,
        count: Number(head.children.find((c) => c instanceof FakeEl && c.hasClass('iplan-tray-count')).textContent),
        collapsed: sec.hasClass('is-collapsed'),
        cards: body.children.filter((c) => c instanceof FakeEl && c.hasClass('iplan-card')).map((c) => c.getAttribute('data-path')),
        note: note ? note.textContent : null,
        connect: note ? (note.children.find((c) => c instanceof FakeEl && c.tagName === 'button') || null) : null,
        headEl: head, secEl: sec,
      };
    });
}
const outlookOf = (root) => sections(root).filter((s) => s.mark.includes('iplan-source-outlook'));
const serial = (root) => JSON.stringify(root.toJSON());
const row = (over) => Object.assign({ ok: true, reason: null, message: null, hint: null, docUrl: null, warning: null, complete: true, count: 1, at: 't1' }, over);

/* ---- 1. the no-change gate ---- */
test('THE GUARD: no outlookAccounts, an empty list, or the default alone: the tray is the single OUTLOOK section it always was, stamped notes included', () => {
  const items = fixtureItems();
  const base = trayFor(ONE_ACCOUNT, items).root;
  const outlook = outlookOf(base);
  assert.equal(outlook.length, 1, 'one Outlook section');
  assert.equal(outlook[0].head, ' OUTLOOK', 'the head is the source label alone, no account');
  assert.equal(outlook[0].count, 20, 'every Outlook note, whatever its source_account, is counted in the one section');
  assert.equal(outlook[0].cards.length, 20);
  assert.equal(outlook[0].note, null);
  // `outlookAccounts: []` and a lone default record (label and all) are the
  // same single account and render the same bytes; so does a list whose
  // second record is unusable and falls out of outlookAccountList.
  const s1 = serial(base);
  assert.equal(serial(trayFor(Object.assign({}, ONE_ACCOUNT, { outlookAccounts: [] }), items).root), s1, 'an empty list changes nothing');
  assert.equal(serial(trayFor(Object.assign({}, ONE_ACCOUNT, { outlookAccounts: [{ accountId: 'default', label: 'Personal' }] }), items).root), s1, 'a lone default record changes nothing, label included');
  assert.equal(serial(trayFor(Object.assign({}, ONE_ACCOUNT, { outlookAccounts: [{ accountId: 'Not An Id' }, 'work'] }), items).root), s1, 'unusable records are not accounts');
  // And the collapse key is still the bare source id.
  assert.equal(outlookOf(trayFor(ONE_ACCOUNT, items, { collapsed: { outlook: true } }).root)[0].collapsed, true);
  assert.equal(outlookOf(trayFor(ONE_ACCOUNT, items, { collapsed: { 'outlook@work': true } }).root)[0].collapsed, false);
  // A per-account row left in the map cannot reach a one-account tray.
  assert.equal(serial(trayFor(ONE_ACCOUNT, items, { byAccount: { default: row({ ok: false, reason: 'unreachable', message: 'down' }) } }).root), s1);
});

/* ---- 2. two accounts, two sections ---- */
test('THE ASK: two accounts, one OUTLOOK section per account, list order, headed by the label (the id when there is none), each counting only its own notes', () => {
  const items = fixtureItems();
  const secs = sections(trayFor(TWO_ACCOUNTS, items).root);
  const outlook = secs.filter((s) => s.mark.includes('iplan-source-outlook'));
  assert.equal(outlook.length, 2, 'two Outlook sections');
  assert.equal(outlook[0].head, ' OUTLOOK · PERSONAL');
  assert.equal(outlook[1].head, ' OUTLOOK · WORK');
  assert.equal(outlook[0].count, 12, 'the notes with no source_account are the default account\'s');
  assert.equal(outlook[1].count, 8, 'the notes stamped work are the work account\'s');
  assert.deepEqual(new Set(outlook[0].cards), new Set(items.filter((i) => i.source === 'outlook' && !i.sourceAccount).map((i) => i.path)));
  assert.deepEqual(new Set(outlook[1].cards), new Set(items.filter((i) => i.sourceAccount === 'work').map((i) => i.path)));
  assert.equal(outlook[0].note, null);
  assert.equal(outlook[1].note, null);
  // The source mark is the source's, on both heads: same glyph, same class.
  assert.equal(outlook[0].mark, outlook[1].mark);
  assert.ok(outlook[0].mark.includes('iplan-source-outlook'));
  // Other sources are untouched: one section each, same head as before.
  const todoist = secs.filter((s) => s.mark.includes('iplan-source-todoist'));
  assert.equal(todoist.length, 1);
  assert.equal(todoist[0].head, ' TODOIST');
  assert.equal(todoist[0].count, 1);
  // The decision the renderer read. One account: one part, keyed and
  // labelled by the source, admitting everything, with no `account` (so the
  // loop takes syncStatus[source] as before). Every other source: the same
  // one part, whatever the account list says.
  const one = T.traySourceSections(ONE_ACCOUNT, 'outlook', items, {});
  assert.equal(one.length, 1);
  assert.equal(one[0].key, 'outlook');
  assert.equal(one[0].label, T.SOURCES.outlook.label);
  assert.equal(one[0].account, null);
  assert.equal(one[0].configured, T.sourceConfigured(ONE_ACCOUNT, 'outlook'));
  assert.ok(items.every((i) => one[0].member(i)));
  for (const src of T.TASK_SOURCES) {
    if (src === 'outlook') continue;
    const p = T.traySourceSections(TWO_ACCOUNTS, src, items, {});
    assert.equal(p.length, 1, `${src}: one section, whatever the account list says`);
    assert.equal(p[0].key, src);
    assert.equal(p[0].label, T.SOURCES[src].label);
    assert.equal(p[0].account, null);
  }
  // Render order follows the list, not the vault, with the default always
  // first (outlookAccountList puts it there).
  const swapped = with2(null, [TWO_ACCOUNTS.outlookAccounts[1], TWO_ACCOUNTS.outlookAccounts[0], { accountId: 'old', label: 'Old' }]);
  assert.deepEqual(outlookOf(trayFor(swapped, items).root).map((s) => s.head), [' OUTLOOK · PERSONAL', ' OUTLOOK · WORK', ' OUTLOOK · OLD']);
  // No label on a record: the id stands in, as everywhere else.
  const unlabelled = with2(null, [{ accountId: 'default' }, { accountId: 'work' }]);
  assert.deepEqual(outlookOf(trayFor(unlabelled, items).root).map((s) => s.head), [' OUTLOOK · DEFAULT', ' OUTLOOK · WORK']);
  // A stamp that spells the default out by hand reads as the default, the
  // way the sync reads it (itemAccountId).
  const typed = fixtureItems().concat([item('outlook', 'typed', { source_account: 'default' })]);
  assert.equal(outlookOf(trayFor(TWO_ACCOUNTS, typed).root)[0].count, 13);
});

/* ---- 3. collapse state is per section ---- */
test('collapsing one account\'s section leaves the other open, and the default keeps its old key', () => {
  const items = fixtureItems();
  assert.equal(T.traySectionKey('outlook', 'default'), 'outlook', 'the reserved default keeps the single-account key');
  assert.equal(T.traySectionKey('outlook', 'work'), 'outlook@work');
  const parts = T.traySourceSections(TWO_ACCOUNTS, 'outlook', items, {});
  assert.deepEqual(parts.map((p) => p.key), ['outlook', 'outlook@work']);
  // A collapse recorded before this build (key `outlook`) still collapses
  // the default's section, and only that one.
  let out = outlookOf(trayFor(TWO_ACCOUNTS, items, { collapsed: { outlook: true } }).root);
  assert.deepEqual(out.map((s) => s.collapsed), [true, false]);
  out = outlookOf(trayFor(TWO_ACCOUNTS, items, { collapsed: { 'outlook@work': true } }).root);
  assert.deepEqual(out.map((s) => s.collapsed), [false, true]);
  // Clicking a head flips exactly its own key.
  const { root, view } = trayFor(TWO_ACCOUNTS, items);
  const live = outlookOf(root);
  live[1].headEl.listeners.click[0]();
  assert.deepEqual(view.collapsed, { 'outlook@work': true });
  assert.equal(live[1].secEl.hasClass('is-collapsed'), true);
  assert.equal(live[0].secEl.hasClass('is-collapsed'), false);
  live[0].headEl.listeners.click[0]();
  assert.deepEqual(view.collapsed, { 'outlook@work': true, outlook: true });
  live[1].headEl.listeners.click[0]();
  assert.deepEqual(view.collapsed, { 'outlook@work': false, outlook: true });
});

/* ---- 4. each section is honest about ITS OWN account ---- */
test('each section reads its own account: not signed in says so under its own head with a Connect button, a failure names its mailbox and shows under it alone, switched off says so and keeps its notes', () => {
  const items = fixtureItems();
  // Work never signed in: "Not connected" under work's head, with the
  // button; the default shows its notes with nothing above them.
  const notYet = with2({ outlookRefreshToken__work: '' });
  let out = outlookOf(trayFor(notYet, items).root);
  assert.equal(out[0].note, null, 'the signed-in default says nothing above its cards');
  assert.equal(out[0].cards.length, 12);
  assert.equal(out[1].note.startsWith(T.TRAY_COPY.unconfigured()), true, 'the work section is unconfigured on its own account');
  assert.ok(out[1].connect, 'and carries the Connect button');
  assert.equal(out[1].connect.getAttribute('aria-label'), 'Connect Outlook · Work');
  assert.equal(out[1].cards.length, 8, 'its notes are still listed under it');
  assert.deepEqual(T.traySourceSections(notYet, 'outlook', items, {}).map((p) => p.configured), [true, false]);
  assert.deepEqual(T.traySourceSections(TWO_ACCOUNTS, 'outlook', items, {}).map((p) => p.configured), [true, true]);
  // The default not signed in, work signed in: the other way round.
  out = outlookOf(trayFor(with2({ outlookRefreshToken: '' }), items).root);
  assert.equal(out[0].note.startsWith(T.TRAY_COPY.unconfigured()), true);
  assert.equal(out[0].connect.getAttribute('aria-label'), 'Connect Outlook · Personal');
  assert.equal(out[1].note, null);

  // After a sync: each section reads its OWN run from syncStatusByAccount,
  // never the folded row. Work unreachable, the default fine: the sentence
  // names work and sits under work's head only, while the folded row (which
  // the board line and the Notice read) leads with the failure.
  const folded = row({ ok: false, reason: 'unreachable', message: 'Work: Outlook is unreachable.', hint: 'Try again.', count: 12 });
  const byAccount = { default: row({ count: 12 }), work: row({ ok: false, reason: 'unreachable', message: 'Work: Outlook is unreachable.', hint: 'Try again.', count: 0 }) };
  out = outlookOf(trayFor(TWO_ACCOUNTS, items, { syncStatus: { outlook: folded }, byAccount }).root);
  assert.equal(out[0].note, null, 'the default is not told about work\'s failure');
  assert.equal(out[1].note, 'Work: Outlook is unreachable.Try again.');
  assert.deepEqual(out.map((s) => s.cards.length), [12, 8], 'the cards stay on screen under the error, as they always did');
  // Nothing for work after a healthy sync of both: "Nothing unscheduled" on
  // ITS count, not the source's.
  const defaultOnly = items.filter((i) => i.sourceAccount !== 'work');
  out = outlookOf(trayFor(TWO_ACCOUNTS, defaultOnly, { syncStatus: { outlook: row({ count: 12 }) }, byAccount: { default: row({ count: 12 }), work: row({ count: 0 }) } }).root);
  assert.equal(out[0].note, null);
  assert.equal(out[1].count, 0);
  assert.equal(out[1].note, T.TRAY_COPY.empty);
  // No sync yet: both wait, whatever a folded row might say.
  out = outlookOf(trayFor(TWO_ACCOUNTS, [], { syncStatus: { outlook: folded } }).root);
  assert.deepEqual(out.map((s) => s.note), [T.TRAY_COPY.unsynced, T.TRAY_COPY.unsynced]);
  // An account added since the last sync has no row yet and waits (its
  // cards, when it has any, show without a note, the rule every section
  // has always had); the other's row is untouched.
  out = outlookOf(trayFor(TWO_ACCOUNTS, defaultOnly, { byAccount: { default: row({ count: 12 }) } }).root);
  assert.deepEqual(out.map((s) => s.note), [null, T.TRAY_COPY.unsynced]);
  out = outlookOf(trayFor(TWO_ACCOUNTS, items, { byAccount: { default: row({ count: 12 }) } }).root);
  assert.deepEqual(out.map((s) => [s.note, s.cards.length]), [[null, 12], [null, 8]]);

  // Switched off: the section stays, its notes stay, and it says what its
  // run would say, prefixed with its label like every other row. A row left
  // in the map from before it was switched off is not read.
  const off = with2(null, [{ accountId: 'default', label: 'Personal' }, { accountId: 'work', label: 'Work', enabled: false }]);
  out = outlookOf(trayFor(off, items, { byAccount: { default: row(), work: row() } }).root);
  assert.equal(out[1].head, ' OUTLOOK · WORK');
  assert.equal(out[1].note, `Work: ${T.OUTLOOK_DISABLED_MESSAGE}`);
  assert.equal(out[1].cards.length, 8, 'switched off deletes nothing and hides nothing');
  assert.equal(out[0].note, null);
  const parts = T.traySourceSections(off, 'outlook', items, { work: row() });
  assert.deepEqual(parts[1].status, { ok: false, reason: 'disabled', message: 'Work: Switched off in the account list.', hint: null, docUrl: null });
  assert.equal(parts[1].configured, true, 'still signed in; switched off is a status, not a missing sign-in');
  // The default switched off: its own run says so (outlookFetchOpen), and
  // the tray says the same without waiting for it.
  const offDefault = with2(null, [{ accountId: 'default', label: 'Personal', enabled: false }, { accountId: 'work', label: 'Work' }]);
  out = outlookOf(trayFor(offDefault, items, {}).root);
  assert.equal(out[0].note, `Personal: ${T.OUTLOOK_DISABLED_MESSAGE}`);
  assert.equal(out[0].cards.length, 12);
  assert.equal(out[1].note, null, 'work has cards and no row yet: the cards, as always');
});

/* ---- 5. an account the list does not name ---- */
test('a note whose stamp nothing lists renders in a trailing section named "(not listed)", never under the first mailbox; an invalid stamp is unlisted too', () => {
  const items = fixtureItems();
  items.push(item('outlook', 'gone-1', { source_account: 'gone' }));
  items.push(item('outlook', 'bad-1', { source_account: 'Not An Id' }));
  const parts = T.traySourceSections(TWO_ACCOUNTS, 'outlook', items, { default: row(), work: row() });
  assert.deepEqual(parts.map((p) => p.key), ['outlook', 'outlook@work', 'outlook@gone', 'outlook@Not An Id']);
  assert.equal(parts[2].label, 'Outlook · gone (not listed)');
  assert.equal(parts[2].configured, false, 'an account nothing lists is never signed in');
  assert.equal(parts[2].status, undefined);
  assert.equal(parts[3].label, 'Outlook · Not An Id (not listed)');
  assert.equal(parts[3].configured, false, 'an invalid id is an account nothing lists: not signed in, and not the default');
  // Even with that id's tokens still in the settings (a record removed, its
  // keys not yet cleared), the section does not claim to be connected.
  const leftover = T.traySourceSections(with2({ outlookRefreshToken__gone: 'rt-g1' }), 'outlook', items, {});
  assert.equal(leftover[2].configured, false);
  const out = outlookOf(trayFor(TWO_ACCOUNTS, items).root);
  assert.deepEqual(out.map((s) => s.head), [' OUTLOOK · PERSONAL', ' OUTLOOK · WORK', ' OUTLOOK · GONE (NOT LISTED)', ' OUTLOOK · NOT AN ID (NOT LISTED)']);
  assert.deepEqual(out.map((s) => s.count), [12, 8, 1, 1]);
  assert.ok(!out[0].cards.includes('02 Planner/outlook/bad-1.md'), 'the first mailbox never shows a note with an invalid source_account');
  assert.ok(!out[0].cards.includes('02 Planner/outlook/gone-1.md'));
  assert.deepEqual(out[2].cards, ['02 Planner/outlook/gone-1.md']);
  assert.deepEqual(out[3].cards, ['02 Planner/outlook/bad-1.md']);
  assert.equal(out[2].note.startsWith(T.TRAY_COPY.unconfigured()), true, 'says it is not connected, with the button to the account list');
  assert.equal(out[2].connect.getAttribute('aria-label'), 'Connect Outlook · gone (not listed)');
  // Every Outlook note is in exactly one section.
  const all = out.flatMap((s) => s.cards);
  assert.equal(all.length, new Set(all).size);
  assert.equal(all.length, items.filter((i) => i.source === 'outlook').length);
  // With one account listed, the same notes stay in the one section (test 1
  // is the rule); nothing is split for a one-account vault.
  assert.equal(outlookOf(trayFor(ONE_ACCOUNT, items).root).length, 1);
});

/* ---- 6. syncNow fills the per-account map ---- */
test('syncNow: with more than one account each run writes its own row, message named for its mailbox; with one account the map stays empty and the row is what it was', async () => {
  function plugin(settings) {
    const p = Object.create(PluginClass.prototype);
    p.settings = settings;
    p.secrets = { mode: 'data-json', available: () => false };
    p.app = { vault: { getAbstractFileByPath: () => null } };
    p._goneProbed = new Set();
    p.syncStatus = {};
    p.syncStatusByAccount = {};
    p.ensureFolders = async () => { };
    p.persistSettings = async () => { };
    p.emitModelChanged = () => { };
    p.connectorDeps = () => ({});
    p.calendarDefsByFeed = {};
    p.calendarFeedSyncedAt = {};
    p.syncing = false;
    return p;
  }
  // Every run answers degraded, so nothing is upserted and no vault is
  // needed: the default is not signed in, work is unreachable.
  const c = T.CONNECTORS.outlook;
  const realFetch = c.fetchOpen;
  c.fetchOpen = async (s) => (s._account === 'work'
    ? T.degraded('outlook', 'unreachable', 'Outlook is unreachable.', 'Try again.')
    : T.degraded('outlook', 'no-token', 'Outlook is not signed in.'));
  const hadWindow = 'window' in globalThis;
  if (!hadWindow) globalThis.window = { setTimeout: () => 0, clearTimeout: () => { } };
  try {
    const p = plugin(Object.assign({}, TWO_ACCOUNTS, { _shadow: {}, calendars: [] }));
    await p.syncNow(false);
    assert.deepEqual(Object.keys(p.syncStatusByAccount), ['default', 'work']);
    const d = p.syncStatusByAccount.default;
    const w = p.syncStatusByAccount.work;
    assert.equal(d.message, 'Personal: Outlook is not signed in.', 'named for its mailbox');
    assert.equal(d.reason, 'no-token');
    assert.equal(w.message, 'Work: Outlook is unreachable.');
    assert.equal(w.hint, 'Try again.');
    assert.equal(w.reason, 'unreachable');
    assert.equal(p.syncStatus.outlook.message, 'Personal: Outlook is not signed in.', 'the folded row is the first failure, as before');
    // The tray then shows each under its own head.
    const out = outlookOf(trayFor(Object.assign({}, TWO_ACCOUNTS, { outlookRefreshToken: 'rt-d1' }), fixtureItems(), { syncStatus: p.syncStatus, byAccount: p.syncStatusByAccount }).root);
    assert.equal(out[0].note, 'Personal: Outlook is not signed in.');
    assert.equal(out[1].note, 'Work: Outlook is unreachable.Try again.');
    // A second sync overwrites each row; an account removed from the list
    // keeps a stale entry that no section reads (test 5: unlisted has none).
    c.fetchOpen = async () => T.okResult('outlook', []);
    await p.syncNow(false);
    assert.equal(p.syncStatusByAccount.default.ok, true);
    assert.equal(p.syncStatusByAccount.work.ok, true);
    // One account, both shapes: nothing written.
    for (const accounts of [undefined, [{ accountId: 'default', label: 'Personal' }]]) {
      const s = Object.assign({}, ONE_ACCOUNT, { _shadow: {}, calendars: [] });
      if (accounts) s.outlookAccounts = accounts;
      const q = plugin(s);
      c.fetchOpen = async () => T.degraded('outlook', 'no-token', 'Outlook is not signed in.');
      await q.syncNow(false);
      assert.deepEqual(q.syncStatusByAccount, {});
      assert.equal(q.syncStatus.outlook.message, 'Outlook is not signed in.', 'no label with one account');
    }
  } finally {
    c.fetchOpen = realFetch;
    if (!hadWindow) delete globalThis.window;
  }
});

/* ---- 7. source scan: the loop reads the part, and nothing else moved ---- */
test('source scan: the by-source loop takes its head label, membership, collapse key, configured and status from the part; the map is written beside the folded row', () => {
  const main = code();
  const start = main.indexOf('/* ---- unscheduled, by source ---- */');
  const end = main.indexOf("const foot = el.createDiv({ cls: 'iplan-tray-foot' });", start);
  assert.ok(start > 0 && end > start);
  const loop = main.slice(start, end);
  assert.match(loop, /for \(const part of traySourceSections\(resolved, key, items, this\.plugin\.syncStatusByAccount\)\) \{/);
  assert.match(loop, /const configured = part\.configured;/);
  assert.match(loop, /const st = part\.account \? part\.status : this\.plugin\.syncStatus\[key\];/, 'a section with an account reads its own row; the single section reads the source\'s, as before');
  assert.match(loop, /i\.source === key && part\.member\(i\) && !i\.plannedDay/);
  assert.match(loop, /headRow\.appendChild\(sourceMarkEl\(key\)\);/, 'the mark is still the source\'s');
  assert.match(loop, /text: ` \$\{part\.label\.toUpperCase\(\)\}`/);
  assert.match(loop, /this\.collapsed\[part\.key\]/);
  assert.doesNotMatch(loop, /this\.collapsed\[key\]/, 'no collapse read on the bare source key remains');
  assert.match(loop, /trayEmptyState\(key, configured, st, list\.length, total, resolved\.secretsInStore === true\)/, 'the empty-state authority is called as before');
  assert.match(loop, /const total = key === MANUAL_SOURCE\s*\n\s*\? items\.filter\(\(i\) => i\.source === MANUAL_SOURCE\)\.length\s*\n\s*: undefined;/, '`total` is still manual\'s alone');
  // syncNow: the folded row's line is untouched and the map is written
  // right after it, only when a label was (more than one account).
  assert.match(main, /this\.syncStatus\[source\] = started\.has\(source\) \? mergeSyncStatus\(this\.syncStatus\[source\], next\) : next;\n\s*started\.add\(source\);/);
  assert.match(main, /if \(label\) this\.syncStatusByAccount\[\(account \|\| accounts\[0\]\)\.accountId\] = next;\n\s*if \(sourceHasAccounts\(source\)\) result\.account/);
  assert.match(main, /this\.syncStatus = \{\};[^\n]*\n\s*this\.syncStatusByAccount = \{\};/, 'initialised beside syncStatus');
  // One sentence for a switched-off account, in the connector and the tray.
  assert.equal(T.OUTLOOK_DISABLED_MESSAGE, 'Switched off in the account list.');
  assert.match(main, /return degraded\('outlook', 'disabled', OUTLOOK_DISABLED_MESSAGE\);/);
  assert.equal((main.match(/Switched off in the account list\./g) || []).length, 1, 'the sentence is written once');
  // The section model asks the registry, never the source id by name.
  const model = main.slice(main.indexOf('function traySectionKey('), main.indexOf('function trayRevealDecision('));
  assert.match(model, /if \(!sourceHasAccounts\(source\)\) return \[whole\];/);
  assert.doesNotMatch(model, /=== 'outlook'/);
  // LF only, as the gate on Windows needs.
  const buf = fs.readFileSync(T.__mainPath);
  let cr = 0;
  for (const b of buf) if (b === 13) cr += 1;
  assert.equal(cr, 0);
});
