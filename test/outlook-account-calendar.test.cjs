/* More than one Microsoft account, part 3b of #38: one Graph calendar feed
 * per account.
 *
 * On sign-in the plugin adds one feed, `{ id: 'outlook-graph', kind:
 * 'graph' }`, and only when no `graph` feed exists, so a second account
 * had no calendar; and every Graph feed was fetched with the one sign-in.
 * Now each `graph` feed may carry an `accountId`; a feed without one is
 * the default's, so the existing `outlook-graph` feed stays with the
 * default account, byte for byte. A further account's feed is
 * `outlook-graph-<id>`, never `outlook-graph`, and is fetched with that
 * account's own token.
 *
 * Gated here, headless:
 *   - THE ASK: a further account's feed is fetched with its own bearer
 *     token, and the default's with the default's; sign-in adds a feed
 *     per account under its own id; the default's one-argument call and
 *     the feed it pushes are what they were;
 *   - ready per account: never signed in, switched off, or unlisted is
 *     not fetched; the row says which; a disabled default keeps its feed
 *     and its sign-in;
 *   - the normaliser carries accountId only for a graph feed that has one;
 *     the key list still gives no slot to a graph feed;
 *   - sign-in and sign-out of a further account add its feed and drop its
 *     events from the board;
 *   - source scan: the connector's ready and fetch go through the feed's
 *     account; the settings row and the url allowlist are untouched.
 *
 * Fixtures are invented: token strings of the shape `at-w1`, a client id
 * of the shape `11111111-...`, labels Personal, Work and Old, no address.
 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const T = require('./harness.cjs');

const PluginClass = require(T.__mainPath);
const raw = () => fs.readFileSync(T.__mainPath, 'utf8');
const code = () => raw().split('\n').filter((l) => !/^\s*(\/\/|\*|\/\*)/.test(l)).join('\n');
const json = (status, body) => ({ status, json: body, text: JSON.stringify(body), headers: {} });
function wire(steps) {
  const calls = [];
  const requestUrl = async (req) => { calls.push(req); const step = steps.shift(); if (!step) throw new Error(`unscripted call: ${req.url}`); return step; };
  return { calls, requestUrl };
}

const CLIENT = '11111111-2222-3333-4444-555555555555';
const FAR = String(Date.now() + 3600000);
const GRAPH_FEED = { id: 'outlook-graph', name: 'Outlook calendar', url: '', color: 2, enabled: true, kind: 'graph' };
const WORK_FEED = { id: 'outlook-graph-work', name: 'Outlook calendar (Work)', url: '', color: 3, enabled: true, kind: 'graph', accountId: 'work' };
const ICS = { id: 'cal-1', name: 'Family', url: 'https://calendar.example.org/private-abc/basic.ics', color: 1, enabled: true, kind: 'ics' };
const THREE = [
  { accountId: 'default', label: 'Personal', enabled: true },
  { accountId: 'work', label: 'Work', enabled: true },
  { accountId: 'old', label: 'Old', enabled: false },
];
const settingsFor = (accounts, extra) => Object.assign({
  outlookClientId: CLIENT, outlookTenant: 'common', outlookScopes: 'Mail.Read Calendars.Read',
  outlookRefreshToken: 'rt-d1', outlookAccessToken: 'at-d1', outlookExpiresAt: FAR, outlookAccount: 'Personal mailbox',
  outlookRefreshToken__work: 'rt-w1', outlookAccessToken__work: 'at-w1', outlookExpiresAt__work: FAR, outlookAccount__work: 'Work mailbox',
  outlookScopes__work: 'Mail.Read Calendars.Read',
  outlookRefreshToken__old: 'rt-o1', outlookAccessToken__old: 'at-o1', outlookExpiresAt__old: FAR, outlookAccount__old: 'Old mailbox',
  outlookScopes__old: 'Mail.Read Calendars.Read',
  calendars: [ICS, GRAPH_FEED, WORK_FEED], _shadow: {},
}, accounts ? { outlookAccounts: accounts } : {}, extra || {});
const row = (id, hh) => ({ id, subject: `E${id}`, isAllDay: false, start: { dateTime: `2026-09-08T${hh}:00:00.0000000`, timeZone: 'UTC' }, end: { dateTime: `2026-09-08T${hh}:30:00.0000000`, timeZone: 'UTC' } });

test('THE ASK (#38, point 1, calendar feeds): each Graph feed is fetched with its own account\'s token, and sign-in adds one feed per account under its own id', async () => {
  const w = T.withSecrets(settingsFor(THREE), null);
  assert.deepEqual(T.enabledCalendarFeeds(w).map((f) => f.id), ['cal-1', 'outlook-graph', 'outlook-graph-work']);
  const icsStub = T.CONNECTORS.calendar.fetchFeed;
  T.CONNECTORS.calendar.fetchFeed = async () => T.okResult('calendar', []);
  try {
    const x = wire([json(200, { value: [row('d', '08')] }), json(200, { value: [row('w', '09')] })]);
    const out = await T.calendarFetchDefs(w, {}, { requestUrl: x.requestUrl, visibleWeekStarts: ['2026-09-07'] });
    assert.equal(x.calls.length, 2);
    assert.deepEqual(x.calls.map((c) => c.headers.Authorization), ['Bearer at-d1', 'Bearer at-w1'], 'the default\'s token for the default\'s feed, work\'s for work\'s');
    for (const c of x.calls) { assert.match(c.url, /^https:\/\/graph\.microsoft\.com\/v1\.0\/me\/calendarView\?/); assert.equal(c.headers.Prefer, 'outlook.timezone="UTC"'); }
    assert.equal(out.ok, true);
    assert.deepEqual(out.items.map((d) => [d.uid, d.feedId, d.feedName]), [['d', 'outlook-graph', 'Outlook calendar'], ['w', 'outlook-graph-work', 'Outlook calendar (Work)']]);
    assert.equal(out.perFeed['outlook-graph-work'].ok, true);
    assert.equal(out.perFeed['outlook-graph'].count, 1);
  } finally { T.CONNECTORS.calendar.fetchFeed = icsStub; }
  // The fetch itself, through the connector's own lambda: the feed's view.
  const y = wire([json(200, { value: [] })]);
  await T.CONNECTORS['outlook-calendar'].fetchFeed(WORK_FEED, w, { requestUrl: y.requestUrl });
  assert.equal(y.calls[0].headers.Authorization, 'Bearer at-w1');
  const z = wire([json(200, { value: [] })]);
  await T.CONNECTORS['outlook-calendar'].fetchFeed(GRAPH_FEED, w, { requestUrl: z.requestUrl });
  assert.equal(z.calls[0].headers.Authorization, 'Bearer at-d1', 'a feed with no accountId is the default\'s');
  // The entry on sign-in, per account. The default's one-argument call and
  // the six-key feed it pushes are what they were; a further account gets
  // its own id, name and accountId, once.
  const settings = { calendars: [ICS] };
  assert.equal(T.ensureGraphCalendarFeed(settings), true);
  assert.deepEqual(settings.calendars[1], { id: 'outlook-graph', name: 'Outlook calendar', url: '', color: 2, enabled: true, kind: 'graph' });
  assert.equal(T.ensureGraphCalendarFeed(settings), false, 'idempotent');
  assert.equal(T.ensureGraphCalendarFeed(settings, 'default'), false, 'the default named is the same call');
  assert.equal(T.ensureGraphCalendarFeed(settings, 'work', 'Work'), true, 'a second mailbox gets a feed of its own');
  assert.deepEqual(settings.calendars[2], { id: 'outlook-graph-work', name: 'Outlook calendar (Work)', url: '', color: 3, enabled: true, kind: 'graph', accountId: 'work' });
  assert.equal(T.ensureGraphCalendarFeed(settings, 'work', 'Work'), false, 'once per account');
  assert.equal(T.ensureGraphCalendarFeed(settings), false, 'and work\'s feed does not stand in for the default\'s');
  assert.equal(settings.calendars.length, 3);
  assert.equal(T.ensureGraphCalendarFeed({ calendars: [WORK_FEED] }), true, 'a vault whose only graph feed is work\'s still gets the default\'s on the default\'s sign-in');
  assert.equal(T.graphFeedId('work'), 'outlook-graph-work');
  assert.notEqual(T.graphFeedId('work'), T.GRAPH_FEED_ID, 'never the default\'s id, or recomputeCalendarDefs merges the two');
  for (const d of ['default', undefined, null, '']) assert.equal(T.graphFeedId(d), 'outlook-graph');
  assert.equal(T.GRAPH_FEED_ID, 'outlook-graph');
  assert.equal(T.graphFeedAccountId(GRAPH_FEED), 'default');
  assert.equal(T.graphFeedAccountId(WORK_FEED), 'work');
  assert.equal(T.graphFeedAccountId({ kind: 'graph', accountId: '' }), 'default');
  assert.equal(T.graphFeedAccountId({ kind: 'graph', accountId: 'Work' }), 'Work', 'as written; it resolves to a disabled account, never to the first mailbox');
  assert.equal(T.calendarFeedFor(w, { feedId: 'outlook-graph-work' }).name, 'Outlook calendar (Work)');
});

test('ready per account: never signed in, switched off or unlisted is not fetched, and the row says which; a disabled default keeps its feed and its sign-in', async () => {
  // Work listed but never signed in: its feed is not ready, the default's is.
  const cold = T.withSecrets(settingsFor(THREE, { outlookRefreshToken__work: '', outlookAccessToken__work: '' }), null);
  assert.deepEqual(T.enabledCalendarFeeds(cold).map((f) => f.id), ['cal-1', 'outlook-graph']);
  assert.equal(T.calendarFeedReady(WORK_FEED, cold), false);
  assert.equal(T.calendarFeedReady(GRAPH_FEED, cold), true);
  assert.equal(T.calendarFeedStatusText(WORK_FEED, null, null, cold), 'Sign in to your Microsoft account under Outlook to connect this calendar.');
  assert.equal(T.calendarFeedStatusText(GRAPH_FEED, null, null, cold), 'Not synced yet this session.');
  assert.equal(T.sourceConfigured(cold, 'outlook-calendar'), true);
  // Its fetch, if forced, is degraded and never a healthy empty calendar.
  const r = await T.CONNECTORS['outlook-calendar'].fetchFeed(WORK_FEED, cold, {});
  assert.equal(r.ok, false);
  assert.equal(r.reason, 'no-token');
  // Old is switched off: its feed is listed, not ready, and says so.
  const oldFeed = { id: 'outlook-graph-old', name: 'Outlook calendar (Old)', url: '', color: 4, enabled: true, kind: 'graph', accountId: 'old' };
  const w = T.withSecrets(settingsFor(THREE, { calendars: [GRAPH_FEED, WORK_FEED, oldFeed] }), null);
  assert.deepEqual(T.enabledCalendarFeeds(w).map((f) => f.id), ['outlook-graph', 'outlook-graph-work']);
  assert.equal(T.outlookFeedReady(oldFeed, w), false);
  assert.equal(T.calendarFeedStatusText(oldFeed, null, null, w), 'Off: its Microsoft account (Old) is switched off in the account list, or not listed.');
  // A feed naming an account nothing lists: not ready, never the default's mailbox.
  const stray = { id: 'outlook-graph-nope', name: 'x', url: '', color: 1, enabled: true, kind: 'graph', accountId: 'nope' };
  assert.equal(T.outlookFeedReady(stray, w), false);
  assert.equal(T.outlookFeedView(w, stray)._account, 'nope');
  assert.equal(T.outlookSignedIn(T.outlookFeedView(w, stray)), false);
  assert.match(T.calendarFeedStatusText(stray, null, null, w), /^Off: its Microsoft account \(nope\)/);
  // The default switched off: its feed stays in the list and in data.json,
  // its sign-in stays, its events are not fetched; work's still are.
  const paused = settingsFor([{ accountId: 'default', label: 'Personal', enabled: false }, { accountId: 'work', label: 'Work' }]);
  const pw = T.withSecrets(paused, null);
  assert.deepEqual(T.enabledCalendarFeeds(pw).map((f) => f.id), ['cal-1', 'outlook-graph-work']);
  assert.equal(T.outlookSignedIn(pw), true, 'still signed in');
  assert.deepEqual(paused.calendars.map((f) => f.id), ['cal-1', 'outlook-graph', 'outlook-graph-work'], 'nothing removed');
  assert.equal(T.calendarFeedStatusText(GRAPH_FEED, null, null, pw), 'Off: its Microsoft account (Personal) is switched off in the account list, or not listed.');
  assert.equal(T.sourceConfigured(pw, 'outlook-calendar'), true, 'work\'s feed keeps the calendar configured');
  // A one-account vault: the pinned readings as they were.
  const one = T.withSecrets(settingsFor(null, { calendars: [ICS, GRAPH_FEED] }), null);
  assert.deepEqual(T.enabledCalendarFeeds(one).map((f) => f.id), ['cal-1', 'outlook-graph']);
  assert.equal(T.calendarFeedStatusText(GRAPH_FEED, null, null, T.withSecrets({ calendars: [GRAPH_FEED], outlookClientId: CLIENT }, null)), 'Sign in to your Microsoft account under Outlook to connect this calendar.');
  assert.equal(T.outlookFeedView(one, GRAPH_FEED), one, 'the default\'s view is the settings themselves');
  assert.equal(T.outlookFeedView(one, WORK_FEED), T.outlookFeedView(one, WORK_FEED), 'one view per account per copy');
});

test('the normaliser carries accountId for a graph feed that has one and for nothing else; the key list gives no slot to any graph feed; the resolver keeps it', () => {
  assert.deepEqual(T.normalizeCalendarFeed(GRAPH_FEED, 1), GRAPH_FEED, 'six keys, as before');
  assert.deepEqual(T.normalizeCalendarFeed(WORK_FEED, 2), WORK_FEED, 'seven, as written');
  assert.deepEqual(Object.keys(T.normalizeCalendarFeed(Object.assign({}, ICS, { accountId: 'work' }), 0)), ['id', 'name', 'url', 'color', 'enabled', 'kind'], 'an iCal feed has no account');
  assert.equal('accountId' in T.normalizeCalendarFeed(Object.assign({}, GRAPH_FEED, { accountId: 7 }), 1), false, 'not a string is no account');
  assert.equal('accountId' in T.normalizeCalendarFeed(Object.assign({}, GRAPH_FEED, { accountId: '' }), 1), false);
  assert.equal(T.normalizeCalendarFeed(Object.assign({}, GRAPH_FEED, { accountId: ' work' }), 1).accountId, ' work', 'as written, never trimmed into a listed account');
  assert.deepEqual(T.calendarFeeds(settingsFor(THREE)).map((f) => f.accountId || null), [null, null, 'work']);
  const w = T.withSecrets(settingsFor(THREE), null);
  assert.equal(w.calendars[2].accountId, 'work', 'the resolved copy carries it');
  const slots = T.secretSlots(settingsFor(THREE));
  assert.ok(!slots.some((s) => /outlook-graph/.test(s.id)), 'no address to keep for a graph feed, whichever account');
  assert.ok(slots.some((s) => s.id === T.calendarSecretKey('cal-1')));
  // A vault with two accounts and the feeds keeps data.json's feed entries as typed.
  assert.deepEqual(T.calendarFeeds({ calendars: [GRAPH_FEED, WORK_FEED] }), [GRAPH_FEED, WORK_FEED]);
});

test('sign-in and sign-out of a further account: its feed is added under its own id on sign-in, its events leave the board on sign-out, the default\'s feed untouched', async () => {
  const s = settingsFor(THREE, { calendars: [GRAPH_FEED], outlookRefreshToken__work: '', outlookAccessToken__work: '' });
  const p = Object.create(PluginClass.prototype);
  Object.assign(p, { settings: s, secrets: null, syncStatus: {}, saved: 0, recomputed: 0, synced: 0, _outlookModal: null });
  p.saveSettings = async () => { p.saved += 1; };
  p.recomputeCalendarDefs = () => { p.recomputed += 1; };
  p.syncNow = () => { p.synced += 1; };
  p.outlookClearPending = () => { };
  const RealNotice = T.__obsidian.Notice;
  T.__obsidian.Notice = function () { };
  try {
    const me = wire([json(200, { userPrincipalName: 'work-mailbox' })]);
    await p.outlookFinishSignIn({ accessToken: 'at-w1', refreshToken: 'rt-w1', expiresIn: 3600, scope: 'Mail.Read Calendars.Read' }, { state: 'st-w', accountId: 'work', scopes: 'a' }, { requestUrl: me.requestUrl });
    assert.deepEqual(s.calendars.map((f) => [f.id, f.accountId || null]), [['outlook-graph', null], ['outlook-graph-work', 'work']]);
    assert.equal(s.calendars[1].name, 'Outlook calendar (Work)');
    assert.equal(s.calendars[1].color, 1, 'the least-used lens, like a pasted feed');
    assert.deepEqual(s.calendars[0], GRAPH_FEED, 'the default\'s feed byte for byte');
    assert.equal(p.synced, 1);
    assert.equal(p.recomputed, 0, 'the sync it started fetches the new feed');
    // Signed in again: no second feed.
    const again = wire([json(200, { userPrincipalName: 'work-mailbox' })]);
    await p.outlookFinishSignIn({ accessToken: 'at-w2', refreshToken: 'rt-w2', expiresIn: 3600, scope: 'a' }, { state: 'st-w2', accountId: 'work', scopes: 'a' }, { requestUrl: again.requestUrl });
    assert.equal(s.calendars.length, 2);
    // Sign out: the feed stays and wants a sign-in; the events leave now.
    await p.outlookSignOut('work');
    assert.equal(s.calendars.length, 2, 'the feed stays, like the default\'s on sign-out');
    assert.equal(p.recomputed, 1);
    assert.equal(T.outlookFeedReady(s.calendars[1], T.withSecrets(s, null)), false);
    assert.equal(T.outlookFeedReady(s.calendars[0], T.withSecrets(s, null)), true, 'the default is still signed in');
  } finally { T.__obsidian.Notice = RealNotice; }
});

test('source scan: the connector\'s ready and fetch go through the feed\'s account; the default\'s calls, the settings row and the url allowlist are untouched', () => {
  const c = code();
  assert.match(c, /ready: \(feed, s\) => outlookFeedReady\(feed, s \|\| \{\}\),/);
  assert.match(c, /fetchFeed: \(feed, s, deps\) => outlookCalendarFetchFeed\(feed, outlookFeedView\(s, feed\), deps\),/);
  assert.match(c, /ready: \(feed\) => !!feedUrl\(feed\)/, 'the iCal connector as it was');
  assert.match(c, /ensureGraphCalendarFeed\(this\.settings\);/, 'the default\'s sign-in makes the literal call');
  assert.match(c, /ensureGraphCalendarFeed\(this\.settings, id, label\);/, 'a further account\'s makes the widened one');
  assert.equal((c.match(/ensureGraphCalendarFeed\(/g) || []).length, 3);
  assert.match(c, /function ensureGraphCalendarFeed\(settings, accountId, label\) \{/);
  assert.match(c, /^    url: typeof f\.url === 'string' \? f\.url\.trim\(\) : '',$/m, 'the one url line on the allowlist, byte for byte');
  assert.match(c, /if \(out\.kind === 'graph' && typeof f\.accountId === 'string' && f\.accountId\) out\.accountId = f\.accountId;/, 'carried as written, graph only');
  assert.match(c, /if \(feed\.kind === 'graph'\) \{[\s\S]{0,400}text: 'Microsoft account'/, 'the row is what it was: the feed\'s name says whose it is');
  assert.match(c, /calendarFeedStatusText\(feed, perFeed\[feed\.id\], secrets, resolved\)/);
  assert.match(c, /const ids = enabledCalendarFeeds\(this\.withSecrets\(\)\)\.map\(\(f\) => f\.id\);/, 'recomputeCalendarDefs drops a feed that is no longer ready, per account now');
  assert.doesNotMatch(c, /'outlook'\s*\]/);
  assert.doesNotMatch(c, /console\.(log|warn|error)\([^)]*(token|Token|clientId)/);
  const cls = c.slice(c.indexOf('async outlookAccountSignOut(id)'), c.indexOf('\n  }\n', c.indexOf('async outlookAccountSignOut(id)')));
  assert.match(cls, /await this\.saveSettings\(\);\n\s*this\.recomputeCalendarDefs\(\);/, 'sign-out drops the events after the save, as the default\'s does');
});
