/* More than one Microsoft account, part 3 of #38: one sync run per account.
 *
 * THE HAZARD (#38, "why this needs agreeing first"): every Outlook note is
 * `source: outlook`, and a sync decides what has finished by source. An
 * open note whose id is missing from a healthy fetch is marked done, then
 * one GET per missing id asks the mailbox whether the message still exists,
 * and a 404 sends the note to the trash. With two mailboxes and notes that
 * do not record their account, a sync of account B treats all of A's notes
 * as missing, B's mailbox answers 404 for A's ids, and A's notes are marked
 * done or trashed, up to 25 per sync.
 *
 * Gated here, headless (no live network, no Obsidian runtime):
 *   - THE ASK: a sync of account B leaves A's notes and the default's
 *     exactly as they are; a disabled account's notes are never touched; a
 *     run that cannot name its account reconciles nothing;
 *   - syncNow: one run per enabled account on its own view, the account
 *     named on every Outlook result, one status row merged; a one-account
 *     vault gets the row it always had;
 *   - the shadow keys and why the separator is `@`;
 *   - every account filter is Outlook's alone: a stray stamp on a Todoist
 *     note changes nothing about Todoist;
 *   - the flag write and the gone probe go to the mailbox the NOTE names;
 *   - one view per account per resolved copy;
 *   - a listed account never signed in answers degraded, never empty;
 *   - Complete on source names every further account lacking the write
 *     permission; a disabled default keeps its feed and its sign-in;
 *   - source scan: the pins the sync core keeps.
 *
 * Fixtures are invented: token strings of the shape `at-w1`, a client id of
 * the shape `11111111-...`, message ids `d1`, `w1`, labels Personal, Work
 * and Old, no address.
 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const T = require('./harness.cjs');

const PluginClass = require(T.__mainPath);
const { TFile, TFolder } = T.__obsidian;
const raw = () => fs.readFileSync(T.__mainPath, 'utf8');
const code = () => raw().split('\n').filter((l) => !/^\s*(\/\/|\*|\/\*)/.test(l)).join('\n');
const json = (status, body) => ({ status, json: body, text: JSON.stringify(body), headers: {} });
function wire(steps) {
  const calls = [];
  const requestUrl = async (req) => { calls.push(req); const step = steps.shift(); if (!step) throw new Error(`unscripted call: ${req.url}`); return step; };
  return { calls, requestUrl };
}

const CLIENT = '11111111-2222-3333-4444-555555555555';
const ROOT = '02 Planner';
const OUTLOOK = `${ROOT}/Outlook`;
const TODOIST = `${ROOT}/Todoist`;
const GRAPH_FEED = { id: 'outlook-graph', name: 'Outlook calendar', url: '', color: 2, enabled: true, kind: 'graph' };
const THREE = [
  { accountId: 'default', label: 'Personal', enabled: true },
  { accountId: 'work', label: 'Work', enabled: true },
  { accountId: 'old', label: 'Old', enabled: false },
];
const FAR = String(Date.now() + 3600000);
// Three sign-ins in the flat namespace (data.json mode, no store): the
// default's bare fields, work's and old's suffixed ones.
const settingsFor = (accounts, extra) => Object.assign({
  plannerFolder: ROOT, outlookClientId: CLIENT, outlookTenant: 'common', outlookScopes: 'Mail.Read Mail.ReadWrite Calendars.Read',
  outlookRefreshToken: 'rt-d1', outlookAccessToken: 'at-d1', outlookExpiresAt: FAR, outlookAccount: 'Personal mailbox',
  outlookRefreshToken__work: 'rt-w1', outlookAccessToken__work: 'at-w1', outlookExpiresAt__work: FAR, outlookAccount__work: 'Work mailbox',
  outlookScopes__work: 'Mail.Read Mail.ReadWrite Calendars.Read',
  outlookRefreshToken__old: 'rt-o1', outlookAccessToken__old: 'at-o1', outlookExpiresAt__old: FAR, outlookAccount__old: 'Old mailbox',
  outlookScopes__old: 'Mail.Read Calendars.Read',
  completeOnSource: true, pushEdits: false, calendars: [GRAPH_FEED], _shadow: {},
}, accounts ? { outlookAccounts: accounts } : {}, extra || {});

function note(id, fm, source) {
  const src = source || 'outlook';
  const f = new TFile();
  f.path = `${src === 'outlook' ? OUTLOOK : TODOIST}/Mail (${src}-${id}).md`;
  f.basename = `Mail (${src}-${id})`;
  f.extension = 'md';
  f.fm = Object.assign({
    type: 'planner-item', source: src, external_id: id, title: `Mail ${id}`,
    status: 'open', due: null, priority: 4, url: null, tags: [], source_status: 'flagged',
    list_id: 'inbox', planned_day: null, planned_half: null, planned_order: 0,
    weekly_goal: false, done_local: false, linked_note: null,
  }, fm || {});
  f.body = '';
  f.stat = { mtime: 1000, ctime: 1000, size: 0 };
  return f;
}
const mail = (id) => ({ id, title: `Mail ${id}`, due: null, priority: 4, description: '', status: 'flagged', listId: 'inbox', tags: [], url: null, parentId: null, recurring: false, dueString: null });
const sh = (over) => Object.assign({ title: null, due: null, priority: 4, description: '', done: false }, over || {});

function plugin(settings, files) {
  const outlook = new TFolder(); outlook.path = OUTLOOK; outlook.children = files.filter((f) => f.path.startsWith(OUTLOOK));
  const todoist = new TFolder(); todoist.path = TODOIST; todoist.children = files.filter((f) => f.path.startsWith(TODOIST));
  const root = new TFolder(); root.path = ROOT; root.children = [outlook, todoist];
  const byPath = new Map(files.map((f) => [f.path, f]));
  const trashed = [];
  const created = [];
  const app = {
    vault: {
      getAbstractFileByPath: (p) => (p === ROOT ? root : (byPath.get(p) || null)),
      cachedRead: async (f) => `---\ntype: planner-item\n---\n${f.body || ''}`,
      create: async (path, content) => { created.push({ path, content }); return null; },
      trash: async (f, system) => {
        assert.equal(system, true);
        trashed.push(f.path);
        for (const dir of [outlook, todoist]) dir.children = dir.children.filter((c) => c !== f);
        byPath.delete(f.path);
      },
    },
    metadataCache: { getFileCache: (f) => ({ frontmatter: f.fm }) },
    fileManager: { processFrontMatter: async (f, fn) => { fn(f.fm); f.stat.mtime += 1; } },
  };
  const p = Object.create(PluginClass.prototype);
  p.settings = settings;
  p.secrets = null;
  p.app = app;
  p._pushTimers = new Map();
  p._syncWrites = new Map();
  p._goneProbed = new Set();
  p.syncStatus = {};
  return { p, trashed, created, byPath };
}
// What syncNow reaches for beyond the sync itself, stubbed.
function headlessSync(p) {
  p.secrets = { mode: 'data-json', available: () => false };
  p.ensureFolders = async () => { };
  p.persistSettings = async () => { };
  p.emitModelChanged = () => { };
  p.connectorDeps = () => ({});
  p.calendarDefsByFeed = {};
  p.calendarFeedSyncedAt = {};
  p.syncing = false;
}
// Records what would leave the machine through the Outlook connector, with
// the account each call was made for, and answers the probe.
function recording(fn, opts) {
  const o = opts || {};
  const c = T.CONNECTORS.outlook;
  const real = { fetchOpen: c.fetchOpen, setClosed: c.setClosed, probeGone: c.probeGone };
  const log = { fetched: [], closed: [], probed: [], notices: [] };
  c.fetchOpen = async (s) => { log.fetched.push(s._account || 'default'); return o.fetch ? o.fetch(s) : T.okResult('outlook', []); };
  c.setClosed = async (s, item, closed) => { log.closed.push([item.id, T.itemAccountId(item), closed]); };
  c.probeGone = async (s, item) => { log.probed.push([item.id, T.itemAccountId(item)]); return o.probe ? o.probe(item) : null; };
  const RealNotice = T.__obsidian.Notice;
  T.__obsidian.Notice = function (msg) { log.notices.push(String(msg)); };
  const hadWindow = 'window' in globalThis;
  if (!hadWindow) globalThis.window = { setTimeout: () => 0, clearTimeout: () => { } };
  return Promise.resolve(fn(log)).finally(() => {
    Object.assign(c, real);
    T.__obsidian.Notice = RealNotice;
    if (!hadWindow) delete globalThis.window;
  });
}
const openFm = (n) => { assert.equal(n.fm.status, 'open', n.path); assert.equal(n.fm.done_local, false, n.path); };

test('THE ASK (#38, point 2): a sync of account B leaves A\'s notes, the default\'s and a disabled account\'s exactly as they are; a run with no account reconciles nothing', async () => {
  const d1 = note('d1'); const d2 = note('d2');
  const w1 = note('w1', { source_account: 'work' }); const w2 = note('w2', { source_account: 'work' });
  const o1 = note('o1', { source_account: 'old' });
  const shadows = { 'outlook:d1': sh(), 'outlook:d2': sh(), 'outlook@work:w1': sh(), 'outlook@work:w2': sh(), 'outlook@old:o1': sh() };
  const { p, trashed, created } = plugin(settingsFor(THREE, { _shadow: shadows }), [d1, d2, w1, w2, o1]);
  // Work's mailbox answers "no such message" to every id it is asked about,
  // which is exactly what B's mailbox says about A's ids.
  await recording(async (log) => {
    await p.upsertSource('outlook', { ok: true, items: [mail('w1'), mail('w3')], account: 'work' });
    assert.deepEqual(trashed, [w2.path], 'only work\'s own vanished note goes');
    assert.deepEqual(log.probed, [['w2', 'work']], 'only work\'s note was asked about');
    for (const n of [d1, d2, w1, o1]) openFm(n);
    assert.ok(p.settings._shadow['outlook:d1'] && p.settings._shadow['outlook:d2'], 'the default\'s shadows are not pruned by work\'s run');
    assert.ok(p.settings._shadow['outlook@old:o1'], 'nor the disabled account\'s');
    assert.equal(p.settings._shadow['outlook@work:w2'], undefined);
    assert.equal(created.length, 1);
    assert.equal(created[0].path, `${OUTLOOK}/Mail w3 (outlook-w3).md`);
    assert.match(created[0].content, /^external_id: "w3"\nsource_account: "work"\ntitle: /m, 'a further account\'s note is stamped, right after its id, in the one write');
    assert.deepEqual(p.settings._shadow['outlook@work:w3'], { due: null, priority: 4, description: '', done: false }, 'under work\'s key');
    // The default's run is symmetric: its mailbox says d2 is gone; work's
    // and old's notes are not its to ask about.
    await p.upsertSource('outlook', { ok: true, items: [mail('d1'), mail('d3')], account: 'default' });
    assert.deepEqual(trashed, [w2.path, d2.path]);
    assert.deepEqual(log.probed, [['w2', 'work'], ['d2', 'default']]);
    for (const n of [d1, w1, o1]) openFm(n);
    assert.equal(created.length, 2);
    assert.doesNotMatch(created[1].content, /source_account/, 'the default\'s note carries no stamp');
    assert.ok(p.settings._shadow['outlook:d3']);
    // A run that cannot name its account: the board still gets its items,
    // and nothing is reconciled, probed, trashed or retried.
    await p.upsertSource('outlook', { ok: true, items: [mail('x9')] });
    assert.deepEqual(trashed, [w2.path, d2.path], 'd1 is absent from this run and is not touched');
    assert.equal(log.probed.length, 2);
    for (const n of [d1, w1, o1]) openFm(n);
    assert.equal(created.length, 3);
    assert.doesNotMatch(created[2].content, /source_account/);
    assert.equal(log.closed.length, 0, 'no flag change went anywhere');
  }, { probe: () => true });
  // The disabled account contributes no run at all, so its notes are never
  // fetched against, whatever its mailbox would say.
  assert.deepEqual(T.outlookExtraRuns(settingsFor(THREE)).map((r) => [r.source, r.account.accountId, r.view._account]), [['outlook', 'work', 'work']]);
  assert.deepEqual(T.outlookExtraRuns(settingsFor(null)), [], 'one account: nothing appended');
  assert.deepEqual(T.outlookExtraRuns(settingsFor([{ accountId: 'default', label: 'Personal' }])), []);
});

test('the reopen retry and the trash of a 404 write stay within the note\'s own account', async () => {
  // Work's note asked to reopen, the default's note asked to reopen: each
  // is retried by its own run only, and a run of the other account never
  // asks its mailbox about it.
  const dr = note('dr', { status: 'done', done_local: false, reopen_pending: true });
  const wr = note('wr', { source_account: 'work', status: 'done', done_local: false, reopen_pending: true });
  const { p } = plugin(settingsFor(THREE, { _shadow: { 'outlook:dr': sh({ done: true }), 'outlook@work:wr': sh({ done: true }) } }), [dr, wr]);
  await recording(async (log) => {
    await p.upsertSource('outlook', { ok: true, items: [mail('w1')], account: 'work' });
    assert.deepEqual(log.probed, [['wr', 'work']]);
    assert.deepEqual(log.closed, [['wr', 'work', false]], 'work\'s reopen goes out on work\'s run');
    assert.equal(p.settings._shadow['outlook@work:wr'].done, false);
    assert.equal(p.settings._shadow['outlook:dr'].done, true, 'the default\'s pending reopen is not work\'s to send');
    await p.upsertSource('outlook', { ok: true, items: [mail('d1')] });
    assert.equal(log.closed.length, 1, 'a run with no account retries nothing');
    await p.upsertSource('outlook', { ok: true, items: [mail('d1')], account: 'default' });
    assert.deepEqual(log.closed, [['wr', 'work', false], ['dr', 'default', false]]);
  }, { probe: () => false });
  // A write that answers "no such message" trashes the note and drops the
  // shadow under the note's own key.
  const wx = note('wx', { source_account: 'work', done_local: true });
  const { p: q, trashed } = plugin(settingsFor(THREE, { _shadow: { 'outlook@work:wx': sh(), 'outlook:wx': sh() } }), [wx]);
  await recording(async () => {
    T.CONNECTORS.outlook.setClosed = async () => { throw T.goneError('Outlook'); };
    await q.detectAndPush(wx.path);
    assert.deepEqual(trashed, [wx.path]);
    assert.equal(q.settings._shadow['outlook@work:wx'], undefined, 'work\'s shadow goes with the note');
    assert.ok(q.settings._shadow['outlook:wx'], 'a default note with the same id keeps its own');
  });
});

test('syncNow: one run per enabled account on its own view, the account named on every Outlook result, one status row; a one-account vault keeps the row it always had', async () => {
  const files = () => [note('d1'), note('d2'), note('w1', { source_account: 'work' }), note('w2', { source_account: 'work' }), note('o1', { source_account: 'old' })];
  const shadows = () => ({ 'outlook:d1': sh(), 'outlook:d2': sh(), 'outlook@work:w1': sh(), 'outlook@work:w2': sh(), 'outlook@old:o1': sh() });
  const fetch = (s) => (s._account === 'work' ? T.okResult('outlook', [mail('w1')], 'more behind the cap', false) : T.okResult('outlook', [mail('d1')]));
  const three = files();
  const { p, trashed } = plugin(settingsFor(THREE, { _shadow: shadows() }), three);
  headlessSync(p);
  await recording(async (log) => {
    await p.syncNow(false);
    assert.deepEqual(log.fetched, ['default', 'work'], 'the default from the registry line, work appended on its own view; old (switched off) never fetched');
    // Each run reconciled its own notes and nobody else's: d2 by the
    // default's run, nothing by work's (its fetch was incomplete), o1 by no one.
    assert.deepEqual(log.probed, [['d2', 'default']]);
    assert.deepEqual(trashed, []);
    assert.equal(three[1].fm.status, 'done', 'd2 left the default\'s open set');
    for (const n of [three[0], three[2], three[3], three[4]]) openFm(n);
    const st = p.syncStatus.outlook;
    assert.equal(st.ok, true);
    assert.equal(st.count, 2, 'the counts add up');
    assert.equal(st.complete, false, 'the source is complete only when every run was');
    assert.equal(st.warning, 'more behind the cap');
    assert.equal(st.message, null);
  }, { fetch });
  // A further account listed but never signed in: its run is degraded and
  // the row says whose failure it is; the default's notes still reconcile.
  const two = files();
  const { p: q } = plugin(settingsFor(THREE, { _shadow: shadows(), outlookRefreshToken__work: '', outlookAccessToken__work: '' }), two);
  headlessSync(q);
  await recording(async (log) => {
    await q.syncNow(false);
    assert.deepEqual(log.fetched, ['default', 'work']);
    const st = q.syncStatus.outlook;
    assert.equal(st.ok, false);
    assert.equal(st.reason, 'no-token');
    assert.equal(st.message, 'Work: Outlook is not signed in.', 'named, since more than one account is listed');
    assert.equal(st.count, 1, 'the default\'s count still shows');
    assert.equal(two[1].fm.status, 'done', 'the default\'s run still reconciled its own');
    openFm(two[3]);
  }, { fetch: (s) => (s._account === 'work' ? T.outlookFetchOpen(s) : T.okResult('outlook', [mail('d1')])) });
  // One account: no run appended, no label, the status row byte for byte
  // what a one-account vault has always had.
  for (const accounts of [null, [{ accountId: 'default', label: 'Personal' }]]) {
    const one = [note('d1'), note('d2')];
    const { p: r } = plugin(settingsFor(accounts, { _shadow: { 'outlook:d1': sh(), 'outlook:d2': sh() } }), one);
    headlessSync(r);
    await recording(async (log) => {
      await r.syncNow(false);
      assert.deepEqual(log.fetched, ['default']);
      assert.deepEqual(r.syncStatus.outlook, { ok: false, reason: 'no-token', message: 'Outlook is not signed in.', hint: null, docUrl: null, warning: null, complete: true, count: 0, at: r.syncStatus.outlook.at });
      for (const n of one) openFm(n);
    }, { fetch: () => T.degraded('outlook', 'no-token', 'Outlook is not signed in.') });
  }
  // The merge itself: a failed run leads, an account switched off never
  // hides a live one, counts add, complete needs every run, first warning.
  const ok = (over) => Object.assign({ ok: true, reason: null, message: null, hint: null, docUrl: null, warning: null, complete: true, count: 1, at: 't1' }, over);
  const m = T.mergeSyncStatus;
  assert.deepEqual(m(null, ok()), ok(), 'nothing prior: the run as it is');
  assert.deepEqual(m(ok({ count: 2 }), ok({ count: 3, at: 't2' })), ok({ count: 5, at: 't2' }));
  assert.deepEqual(m(ok({ warning: 'w1' }), ok({ warning: 'w2', at: 't2', complete: false })), ok({ warning: 'w1', at: 't2', complete: false, count: 2 }));
  const bad = ok({ ok: false, reason: 'unreachable', message: 'Work: down', count: 0 });
  assert.equal(m(ok(), bad).message, 'Work: down', 'the failed run leads whichever order');
  assert.equal(m(bad, ok({ at: 't2' })).ok, false);
  assert.equal(m(bad, ok({ at: 't2' })).count, 1);
  const off = ok({ ok: false, reason: 'disabled', message: 'Switched off in the account list.', count: 0 });
  assert.equal(m(off, ok({ at: 't2' })).ok, true, 'switched off is not a failure');
  assert.equal(m(ok(), off).ok, true);
  assert.equal(m(off, off).ok, false, 'both switched off: the row says so');
});

test('the shadow keys: the default and every other source keep source:id; a further account is source@account:id, because the prune scans by the source: prefix', () => {
  assert.equal(T.shadowKey('outlook', 'default', 'x'), 'outlook:x');
  assert.equal(T.shadowKey('outlook', undefined, 'x'), 'outlook:x');
  assert.equal(T.shadowKey('outlook', '', 'x'), 'outlook:x');
  assert.equal(T.shadowKey('todoist', null, '7'), 'todoist:7', 'every shadow already in data.json keeps its key');
  assert.equal(T.shadowKey('outlook', 'work', 'x'), 'outlook@work:x');
  assert.equal(T.shadowPrefix('outlook', 'default'), 'outlook:');
  assert.equal(T.shadowPrefix('outlook', 'work'), 'outlook@work:');
  assert.equal(T.sourceHasAccounts('outlook'), true);
  for (const s of ['todoist', 'clickup', 'email', 'manual', 'calendar', undefined]) assert.equal(T.sourceHasAccounts(s), false, String(s));
  // The account a shadow is keyed under: the stamp for Outlook, the default
  // for every other source whatever a stray stamp says.
  assert.equal(T.itemShadowAccount({ source: 'outlook', sourceAccount: 'work' }), 'work');
  assert.equal(T.itemShadowAccount({ source: 'outlook' }), 'default');
  assert.equal(T.itemShadowAccount({ source: 'todoist', sourceAccount: 'work' }), 'default');
  assert.equal(T.itemShadowAccount(null), 'default');
  // THE COLLISION the issue's `:` shape would have: the default's prune
  // selects `outlook:` keys and reads the rest as the id. Under
  // `outlook:work:x` it would read `work:x`, an id in neither of its sets,
  // and drop account B's done shadow (the record that lets an uncheck
  // reopen a task) on every default sync. Under `@` it never sees it.
  const now = 10;
  const map = { 'outlook:a': sh(), 'outlook:work:x': sh({ done: true, doneAt: now }), 'outlook@work:x': sh({ done: true, doneAt: now }), 'outlook@work:gone': sh() };
  assert.deepEqual(T.pruneShadows(map, 'outlook', new Set(['a']), new Set(['a']), now), ['outlook:work:x'], 'the default\'s prune (the 5-arg call as it ships) would eat the colon shape and leaves the @ shape');
  assert.deepEqual(T.pruneShadows(map, 'outlook', new Set(['x']), new Set(['x']), now, undefined, 'work'), ['outlook@work:gone'], 'work\'s prune sees work\'s keys only');
  assert.deepEqual(T.pruneShadows(map, 'outlook', new Set(['x']), new Set(['x']), now, undefined, 'default'), ['outlook:a', 'outlook:work:x'], 'default named is the bare prefix');
  // The item key the subtask index uses is untouched: no source with
  // accounts has subtasks, and a Todoist parent is found as before.
  const idx = T.buildItemIndex([{ source: 'todoist', id: 'p', parentId: null }, { source: 'todoist', id: 'c', parentId: 'p', sourceAccount: 'work' }]);
  assert.equal((idx.childrenOf.get('todoist:p') || []).length, 1);
});

test('every account filter is Outlook\'s alone: a stray source_account on a Todoist note changes nothing about Todoist', async () => {
  const t1 = note('t1', { source_account: 'work' }, 'todoist');
  const t2 = note('t2', {}, 'todoist');
  const { p, trashed } = plugin(settingsFor(THREE, { todoistToken: 'tok', _shadow: { 'todoist:t1': sh(), 'todoist:t2': sh() } }), [t1, t2]);
  const c = T.CONNECTORS.todoist;
  const real = { probeGone: c.probeGone, setClosed: c.setClosed };
  const closed = [];
  c.probeGone = async () => null;
  c.setClosed = async (s, item, done) => { closed.push([item.id, done]); };
  const RealNotice = T.__obsidian.Notice;
  T.__obsidian.Notice = function () { };
  const hadWindow = 'window' in globalThis;
  if (!hadWindow) globalThis.window = { setTimeout: () => 0, clearTimeout: () => { } };
  try {
    await p.upsertSource('todoist', { ok: true, items: [{ id: 't2', title: 'Task t2', due: null, priority: 4, description: '', status: null, listId: null, tags: [], url: null, parentId: null, recurring: null, dueString: null }] });
    assert.equal(t1.fm.status, 'done', 'absent from Todoist\'s open set, so done: the stamp did not drop it out of Todoist\'s reconcile');
    assert.equal(t1.fm.done_local, true);
    assert.deepEqual(trashed, []);
    assert.equal(p.settings._shadow['todoist:t1'].done, true, 'under the bare key');
    // The edit-time path reads the same bare key: a check on the stamped
    // note reaches Todoist.
    t1.fm.status = 'open'; t1.fm.done_local = true; p.settings._shadow['todoist:t1'].done = false;
    p.clearSyncWrite(t1.path); // the reconcile above announced its own write; this check is the person's
    await p.detectAndPush(t1.path);
    assert.deepEqual(closed, [['t1', true]]);
    assert.equal(p.settings._shadow['todoist:t1'].done, true);
    await p.removeGoneItem('todoist', T.itemFromFrontmatter(t1.fm, t1.path, t1.basename));
    assert.equal(p.settings._shadow['todoist:t1'], undefined, 'and the gone path deletes the bare key');
  } finally {
    Object.assign(c, real);
    T.__obsidian.Notice = RealNotice;
    if (!hadWindow) delete globalThis.window;
  }
});

test('the flag write and the gone probe go to the mailbox the NOTE names, through the connector; a stamp nothing lists, or an account switched off, gets no call', async () => {
  const w = T.withSecrets(settingsFor(THREE), null);
  const item = (id, account) => Object.assign({ id, source: 'outlook' }, account ? { sourceAccount: account } : {});
  // The write: work's token for work's note, the default's for an unstamped one.
  let x = wire([json(200, {})]);
  await T.outlookSetClosed(w, item('w1', 'work'), true, { requestUrl: x.requestUrl });
  assert.equal(x.calls[0].headers.Authorization, 'Bearer at-w1', 'work\'s bearer token');
  assert.match(x.calls[0].url, /\/me\/messages\/w1$/);
  assert.equal(x.calls[0].method, 'PATCH');
  x = wire([json(200, {})]);
  await T.outlookSetClosed(w, item('d1'), false, { requestUrl: x.requestUrl });
  assert.equal(x.calls[0].headers.Authorization, 'Bearer at-d1', 'the default\'s bearer token');
  // Refused before any call: switched off, unlisted, an id the rule refuses,
  // and the default's own guards as they were.
  for (const [it, re] of [[item('o1', 'old'), /switched off/], [item('n1', 'nope'), /switched off|not listed/], [item('W1', 'Work'), /switched off|not listed/]]) {
    x = wire([]);
    await assert.rejects(T.outlookSetClosed(w, it, true, { requestUrl: x.requestUrl }), re);
    assert.equal(x.calls.length, 0);
  }
  await assert.rejects(T.outlookSetClosed({ outlookClientId: CLIENT }, item('d1'), true, {}), /not signed in/);
  const readOnly = T.withSecrets(settingsFor(THREE, { outlookScopes__work: 'Mail.Read' }), null);
  await assert.rejects(T.outlookSetClosed(readOnly, item('w1', 'work'), true, {}), /Mail\.ReadWrite/, 'the write permission is per account');
  // The probe: the same resolution; no evidence for an account that cannot be asked.
  x = wire([json(200, { id: 'w2' })]);
  assert.equal(await T.outlookProbeGone(w, item('w2', 'work'), { requestUrl: x.requestUrl }), false);
  assert.equal(x.calls[0].headers.Authorization, 'Bearer at-w1');
  assert.match(x.calls[0].url, /\/me\/messages\/w2\?\$select=id$/);
  x = wire([json(404, { error: { code: 'ErrorItemNotFound', message: 'not found' } })]);
  assert.equal(await T.outlookProbeGone(w, item('w2', 'work'), { requestUrl: x.requestUrl }), true, 'work\'s mailbox says gone');
  for (const it of [item('o1', 'old'), item('n1', 'nope'), item('W1', 'Work')]) {
    x = wire([]);
    assert.equal(await T.outlookProbeGone(w, it, { requestUrl: x.requestUrl }), null, JSON.stringify(it));
    assert.equal(x.calls.length, 0);
  }
  assert.equal(await T.outlookProbeGone({ outlookClientId: CLIENT }, item('d1'), {}), null, 'not signed in: as it was');
  // The plugin's own path hands the connector the note, never an account.
  const { p } = plugin(settingsFor(THREE), []);
  await recording(async (log) => {
    await p.applyDoneOnSource(T.itemFromFrontmatter(note('w1', { source_account: 'work' }).fm, 'p', 'b'), true);
    assert.deepEqual(log.closed, [['w1', 'work', true]]);
  });
  assert.match(code(), /await c\.setClosed\(this\.withSecrets\(\), item, closed\);/, 'the pinned call site did not move');
  // One view per account per resolved copy: a token rotated inside one
  // call is what the next call on the same copy reads. The live settings
  // are never a cache.
  const live = settingsFor(THREE);
  const copy = T.withSecrets(live, null);
  assert.equal(T.outlookAccountView(copy, 'work'), T.outlookAccountView(copy, 'work'), 'the same view twice on one copy');
  assert.notEqual(T.outlookAccountView(live, 'work'), T.outlookAccountView(live, 'work'), 'no cache on the live settings');
  assert.equal(Object.getOwnPropertyDescriptor(copy, '_accountViews').enumerable, false, 'the cache is hidden like the links');
  assert.doesNotMatch(JSON.stringify(copy), /_accountViews/, 'and never reaches a further copy or disk through the copy');
  const y = wire([json(200, { access_token: 'at-w2', refresh_token: 'rt-w2', expires_in: 3600 })]);
  await T.ensureAccessToken(T.outlookAccountView(copy, 'work'), { requestUrl: y.requestUrl, now: () => 5000 }, true);
  assert.equal(T.outlookAccountView(copy, 'work').outlookAccessToken, 'at-w2', 'the second view on the copy reads the rotated token');
  assert.equal(T.outlookAccountView(copy, 'work').outlookRefreshToken, 'rt-w2');
  assert.equal(copy.outlookAccessToken, 'at-d1', 'and the default\'s did not move');
  assert.equal(live.outlookAccessToken__work, 'at-w2', 'the rotation reached the live settings');
  assert.equal(T.outlookItemAccount(copy, item('w1', 'work')).view, T.outlookAccountView(copy, 'work'));
  assert.equal(T.outlookItemAccount(copy, item('d1')).view, copy, 'the default is the copy itself');
  assert.equal(T.outlookItemAccount(copy, item('o1', 'old')).account.enabled, false);
});

test('a listed account never signed in answers degraded, never a healthy empty set; an account switched off answers switched off; the default switched off keeps its feed and its sign-in', async () => {
  const w = T.withSecrets(settingsFor(THREE, { outlookRefreshToken__work: '', outlookAccessToken__work: '' }), null);
  const r = await T.outlookFetchOpen(T.outlookAccountView(w, 'work'));
  assert.equal(r.ok, false, 'never `ok` with `items: []`, which would retire every note of the account');
  assert.equal(r.reason, 'no-token');
  assert.deepEqual(r.items, []);
  const off = await T.outlookFetchOpen(T.outlookAccountView(w, 'old'));
  assert.equal(off.ok, false);
  assert.equal(off.reason, 'disabled');
  assert.equal(off.message, 'Switched off in the account list.');
  // The default switched off: its run is degraded, and nothing else about
  // it moves (no sign-out, no feed removed). The pinned guards as they were.
  const paused = settingsFor([{ accountId: 'default', label: 'Personal', enabled: false }, { accountId: 'work', label: 'Work' }]);
  const pw = T.withSecrets(paused, null);
  const d = await T.outlookFetchOpen(pw);
  assert.equal(d.reason, 'disabled');
  assert.equal(T.outlookSignedIn(pw), true, 'still signed in');
  assert.deepEqual(paused.calendars, [GRAPH_FEED], 'the outlook-graph feed stays');
  assert.deepEqual(T.enabledCalendarFeeds(pw).map((f) => f.id), ['outlook-graph'], 'and is still read in this part; the calendar per account is the next one');
  assert.equal((await T.outlookFetchOpen({})).reason, 'no-token');
  assert.equal((await T.outlookFetchOpen({ outlookClientId: CLIENT })).reason, 'no-token');
  assert.equal((await T.outlookFetchOpen({ outlookAccounts: THREE })).reason, 'no-token', 'the flat settings are the default, enabled');
});

test('Complete on source: every further account lacking the write permission is named and re-consented one at a time, never several sign-ins at once', () => {
  // work signed in read-only, old switched off but signed in read-only, and
  // a fourth listed and never signed in.
  const s = settingsFor(THREE.concat([{ accountId: 'fresh', label: 'Fresh' }]), { outlookScopes__work: 'Mail.Read Calendars.Read' });
  const w = T.withSecrets(s, null);
  assert.deepEqual(T.outlookAccountsNeedingWrite(w).map((a) => a.accountId), ['work', 'old'], 'signed in without Mail.ReadWrite; never the default (its own line handles it), never one that has not signed in');
  assert.deepEqual(T.outlookAccountsNeedingWrite(T.withSecrets(settingsFor(THREE), null)).map((a) => a.accountId), ['old']);
  assert.deepEqual(T.outlookAccountsNeedingWrite(T.withSecrets(settingsFor(null), null)), [], 'one account: nothing to name');
  assert.equal(T.outlookWriteConsentNotice(T.outlookAccountsNeedingWrite(w)), 'Planner: Work, Old need the Mail.ReadWrite permission too. Use "Sign in again" on each row under Outlook, one at a time.');
  assert.equal(T.outlookWriteConsentNotice([{ label: 'Work' }]), 'Planner: Work needs the Mail.ReadWrite permission too. Use "Sign in again" on its row under Outlook, one at a time.');
  const c = code();
  assert.match(c, /if \(v && outlookSignedIn\(r\) && !outlookHasWriteScope\(r\)\) this\.plugin\.outlookSignIn\(\{ write: true, onDone: \(\) => this\.display\(\) \}\);\n\s*const more = v \? outlookAccountsNeedingWrite\(r\) : \[\];\n\s*if \(more\.length\) new Notice\(outlookWriteConsentNotice\(more\), 12000\);/,
    'the default\'s pinned line, then the further accounts are named; only one sign-in is ever started');
  assert.equal((c.match(/outlookSignIn\(\{ write: true/g) || []).length, 1, 'no second browser sign-in is started from the toggle');
  // The row's status line is what says "sign in again" for a further account, per account.
  assert.match(T.outlookStatusText(T.outlookAccountView(w, 'work')), /Signed in as Work mailbox\. Flag changes need one more permission/);
  assert.equal(T.outlookStatusText(T.outlookAccountView(T.withSecrets(settingsFor(THREE), null), 'work')), 'Signed in as Work mailbox.');
});

test('source scan: the pins the sync core keeps, and every shadow key in the class goes through shadowKey', () => {
  const c = code();
  const cls = c.slice(c.indexOf('class IcorPlannerPlugin'), c.indexOf('class IcorPlannerSettingTab'));
  assert.equal((c.match(/const s = this\.withSecrets\(\);/g) || []).length, 4);
  assert.match(c, /await this\.ensureFolders\(\);\n\s*if \(this\.secrets\.mode === 'env-file'\) await this\.envStore\.load\(\);\n\s*const s = this\.withSecrets\(\);\n\s*const runs = SYNCED_SOURCES\.map\(\(k\) => \[k, CONNECTORS\[k\]\.fetchOpen\(s\)\]\);\n\s*for \(const r of outlookExtraRuns\(s\)\) runs\.push\(/, 'the extra runs are appended after the pinned registry line');
  assert.match(c, /if \(sourceHasAccounts\(source\)\) result\.account = account \? account\.accountId : OUTLOOK_DEFAULT_ACCOUNT;\n\s*if \(result\.ok\) await this\.upsertSource\(source, result\);/, 'every Outlook run names its account on its result, the default included');
  assert.match(c, /this\.syncStatus\[source\] = started\.has\(source\) \? mergeSyncStatus\(this\.syncStatus\[source\], next\) : next;/, 'the merge runs only for a second run of one source');
  assert.match(c, /const accountMissing = hasAccounts && !accountId;/);
  assert.match(c, /&& ownsNote\(it\) && !accountMissing\)/, 'the account guard sits inside the pinned reconcile predicate');
  assert.match(c, /if \(accountMissing \|\| !ownsNote\(it\)\) continue;/, 'and in the reopen retry');
  assert.match(c, /else await this\.createItemFile\(folder, source, t\);/, 'the default\'s create is the literal call');
  assert.match(c, /await this\.createItemFile\(folder, source, t, accountId\);/);
  assert.match(c, /pruneShadows\(s\._shadow, source, new Set\(existing\.keys\(\)\), openIds, nowMs, undefined, accountId\)/);
  for (const fn of ['async removeGoneItem(source, item) {\n    const key = shadowKey(source, itemShadowAccount(item), item.id);', 'const key = shadowKey(source, itemShadowAccount(it), it.id);\n      if (this._goneProbed.has(key)) continue;', 'const key = shadowKey(item.source, itemShadowAccount(item), item.id);\n    const sh = s._shadow[key];']) {
    assert.ok(cls.includes(fn), fn);
  }
  assert.doesNotMatch(cls, /`\$\{(source|item\.source)\}:\$\{/, 'no bare source:id key is built in the class any more');
  assert.doesNotMatch(c, /source === 'outlook'/, 'the source with accounts is asked from the registry, not by name');
  assert.doesNotMatch(c, /'outlook'\s*\]/);
  assert.equal((c.match(/ensureGraphCalendarFeed\(/g) || []).length, 2, 'no calendar feed for a further account in this part');
  assert.doesNotMatch(c, /later update/, 'the sign-in row and notice no longer defer the sync');
  assert.match(c, /new Notice\(`Planner: signed in to Outlook \(\$\{label\}\)\$\{account \? ` as \$\{account\}` : ''\}\.`\);\n\s*if \(pending\.onDone\) pending\.onDone\(\);\n\s*this\.syncNow\(false\);/, 'a further account syncs on sign-in, like the default');
  assert.match(c, /function outlookFetchOpen\(settings, deps\) \{\n\s*const s = settings \|\| \{\};\n\s*if \(!outlookAccountById\(s, s\._account\)\.enabled\) return degraded\('outlook', 'disabled'/);
  // The stamp: one writer, guarded, and the note reader unchanged.
  assert.match(c, /\.\.\.\(accountId && accountId !== OUTLOOK_DEFAULT_ACCOUNT \? \[`source_account: \$\{JSON\.stringify\(String\(accountId\)\)\}`\] : \[\]\),/);
  assert.equal(c.split('\n').filter((l) => l.includes('source_account')).length, 2, 'the reader and the writer, no third line');
  assert.doesNotMatch(c, /console\.(log|warn|error)\([^)]*(token|Token|clientId)/);
});
