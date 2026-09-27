/* More than one Microsoft account, part 2 of #38: one sign-in and one set
 * of store keys per account.
 *
 * `default` keeps the flat outlook* fields and the four store keys members
 * already have. A further account's values live in the same flat settings
 * namespace under a suffixed field (`outlookRefreshToken__work`); its store
 * key carries the id as a suffix (`icor-for-life-planner-outlook-refresh-
 * token-work`, env key OUTLOOK_REFRESH_TOKEN_WORK). Nothing syncs a further
 * account yet.
 *
 * Gated here, pure and headless (no live network, no Obsidian runtime):
 *   - THE ASK (#38, point 4): no `outlookAccounts` key, or `default` only:
 *     the walked field list is SECRET_FIELDS, the key list is what it was,
 *     data.json loads and saves byte for byte;
 *   - the field and key naming, and its strict inverse;
 *   - the six walkers see a second token; an unlisted account is walked;
 *   - the view: the default IS the settings object; a rotation through a
 *     further account's view lands on its own keys;
 *   - the pending map keyed by state: a reply finds its own sign-in;
 *   - a further account's sign-in and sign-out leave the default's keys;
 *   - source scan: the pins the sign-in must keep, and the settings rows.
 *
 * Fixtures are invented: token strings of the shape `rt-w1`, a client id of
 * the shape `11111111-...`, labels Personal and Work, no address.
 */
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const T = require('./harness.cjs');

const PluginClass = require(T.__mainPath);
const raw = () => fs.readFileSync(T.__mainPath, 'utf8');
const code = () => raw().split('\n').filter((l) => !/^\s*(\/\/|\*|\/\*)/.test(l)).join('\n');

class FakeSecretStorage {
  constructor() { this.m = new Map(); }
  setSecret(id, secret) { if (!/^[a-z0-9-]+$/.test(id)) throw new Error(`invalid secret id: ${id}`); this.m.set(id, String(secret)); }
  getSecret(id) { return this.m.has(id) ? this.m.get(id) : null; }
  listSecrets() { return [...this.m.keys()]; }
}
const store = () => { const storage = new FakeSecretStorage(); return { storage, vault: new T.SecretVault(storage) }; };
function wire(steps) {
  const calls = [];
  const requestUrl = async (req) => { calls.push(req); const step = steps.shift(); if (!step) throw new Error(`unscripted call: ${req.url}`); return step; };
  return { calls, requestUrl };
}
const json = (status, body) => ({ status, json: body, text: JSON.stringify(body), headers: {} });
const form = (req) => Object.fromEntries(new URLSearchParams(req.body));

const CLIENT = '11111111-2222-3333-4444-555555555555';
const PREFIX = 'icor-for-life-planner-';
const GRAPH_FEED = { id: 'outlook-graph', name: 'Outlook calendar', url: '', color: 2, enabled: true, kind: 'graph' };
const DEFAULT_ONLY = [{ accountId: 'default', label: 'Personal', enabled: true }];
const TWO = [{ accountId: 'default', label: 'Personal', enabled: true }, { accountId: 'work', label: 'Work', enabled: true }];
const SEVEN = ['todoistToken', 'clickupToken', 'imapPassword', 'outlookRefreshToken', 'outlookAccessToken', 'outlookExpiresAt', 'outlookAccount'];
const FOUR = SEVEN.slice(3);
const WORK_FOUR = FOUR.map((f) => `${f}__work`);
const DEFAULT_KEYS = ['outlook-refresh-token', 'outlook-access-token', 'outlook-expires-at', 'outlook-account'].map((k) => PREFIX + k);
const WORK_KEYS = DEFAULT_KEYS.map((k) => `${k}-work`);

// Both accounts signed in, the values still in the settings.
const twoSignedIn = () => Object.assign({}, T.DEFAULT_SETTINGS, {
  outlookClientId: CLIENT, outlookScopes: 'Mail.Read Calendars.Read', calendars: [GRAPH_FEED], outlookAccounts: TWO,
  outlookRefreshToken: 'rt-d1', outlookAccessToken: 'at-d1', outlookExpiresAt: '1000', outlookAccount: 'Personal mailbox',
  outlookRefreshToken__work: 'rt-w1', outlookAccessToken__work: 'at-w1', outlookExpiresAt__work: '2000', outlookAccount__work: 'Work mailbox',
  outlookScopes__work: 'Mail.Read Mail.ReadWrite Calendars.Read',
});
// The plugin without Obsidian: the hooks the sign-in methods reach for.
function headless(settings, vault) {
  const p = Object.create(PluginClass.prototype);
  Object.assign(p, { settings, secrets: vault, syncStatus: {}, syncStatusByAccount: {}, saved: [], recomputed: 0, synced: 0 });
  p.saveSettings = async () => { p.saved.push(JSON.stringify(p.settings)); };
  p.recomputeCalendarDefs = () => { p.recomputed += 1; };
  p.syncNow = () => { p.synced += 1; };
  return p;
}
// The PR 1 round trip: what a 0.16.2 data.json holds, through load and save.
async function loadAndSave(extra) {
  const text = JSON.stringify(Object.assign({}, T.DEFAULT_SETTINGS, { outlookClientId: CLIENT, calendars: [GRAPH_FEED], secretsInStore: true }, extra || {}, { _shadow: {} }), null, 2);
  const p = Object.create(PluginClass.prototype);
  p.secretStorage = null;
  p.secrets = p.vaultFor('secret-storage');
  const adopted = T.adoptSettings(JSON.parse(text), p.secrets);
  p.settings = adopted.settings;
  p.saveData = async (s) => { p.saved = JSON.stringify(s, null, 2); };
  await p.persistSettings();
  assert.equal(adopted.changed, false, 'no write-back at load');
  assert.equal(p.saved, text, 'byte for byte at save');
}

test('THE ASK (#38, point 4): with no outlookAccounts key, or default only, the walked fields, the key list and data.json are what they were', async () => {
  assert.deepEqual(Object.keys(T.SECRET_FIELDS), SEVEN, 'the table members already have keys under is frozen');
  for (const s of [{}, null, { outlookAccounts: DEFAULT_ONLY }, Object.assign({}, T.DEFAULT_SETTINGS, { outlookRefreshToken: 'rt-d1' })]) {
    assert.deepEqual(T.secretFieldNames(s), SEVEN, JSON.stringify(s));
    assert.deepEqual(T.outlookSecretAccountIds(s), []);
  }
  const keys = (s) => T.secretSlots(s).map((slot) => [slot.id, slot.label, slot.envKey, slot.ids]);
  assert.deepEqual(keys({ calendars: [GRAPH_FEED], outlookAccounts: DEFAULT_ONLY }), keys({ calendars: [GRAPH_FEED] }));
  assert.equal(T.secretSlots({}).length, 5);
  await loadAndSave();
  await loadAndSave({ outlookAccounts: DEFAULT_ONLY });
});

test('the naming: the default keeps the bare field, a further account gets a suffix, the inverse is strict, the key follows', () => {
  assert.deepEqual(T.OUTLOOK_ACCOUNT_SECRET_FIELDS, FOUR);
  assert.deepEqual(T.OUTLOOK_ACCOUNT_FIELDS, FOUR.concat(['outlookScopes']));
  for (const f of T.OUTLOOK_ACCOUNT_FIELDS) {
    for (const d of ['default', undefined, '']) assert.equal(T.outlookAccountField(d, f), f);
    assert.equal(T.outlookAccountField('work', f), `${f}__work`);
  }
  assert.throws(() => T.outlookAccountField('Work', 'outlookRefreshToken'), /not an account id/, 'never folded, never fixed up');
  assert.throws(() => T.outlookAccountField('work', 'outlookClientId'), /not an account field/, 'the client id is shared; no per-account field');
  assert.deepEqual(T.outlookAccountFieldParts('outlookRefreshToken__work'), { field: 'outlookRefreshToken', accountId: 'work' });
  assert.deepEqual(T.outlookAccountFieldParts('outlookAccount__work-2'), { field: 'outlookAccount', accountId: 'work-2' });
  for (const bad of ['outlookRefreshToken', 'outlookRefreshToken__', 'outlookRefreshToken__default', 'outlookRefreshToken__Work', 'outlookRefreshToken__a--b', 'outlookScopes__work', 'todoistToken__work', '__work', null]) {
    assert.equal(T.outlookAccountFieldParts(bad), null, JSON.stringify(bad));
  }
  assert.deepEqual(WORK_FOUR.map(T.fieldSecretKey), WORK_KEYS);
  assert.equal(T.fieldSecretKey('outlookAccount__work-2'), `${PREFIX}outlook-account-work-2`);
  assert.equal(T.envKeyFor(T.fieldSecretKey('outlookRefreshToken__work')), 'OUTLOOK_REFRESH_TOKEN_WORK');
  assert.equal(T.envKeyFor(T.fieldSecretKey('outlookAccount__work-2')), 'OUTLOOK_ACCOUNT_WORK_2');
  assert.deepEqual(FOUR.map(T.fieldSecretKey), DEFAULT_KEYS, 'the default keys are what they were');
  for (const bad of ['outlookRefreshToken__Work', 'outlookScopes__work', 'imapUser']) assert.throws(() => T.fieldSecretKey(bad), /not a secret field/, bad);
  assert.deepEqual(T.secretFieldNames({ outlookAccounts: TWO }), SEVEN.concat(WORK_FOUR));
  assert.deepEqual(T.outlookSecretAccountIds({ outlookAccounts: TWO, outlookRefreshToken__old: 'rt-o1', outlookAccount__zed: 'x', outlookRefreshToken__blank: ' ' }), ['work', 'old', 'zed'], 'listed first, then a value under a suffixed field whose record is gone; a blank is not an account');
  assert.deepEqual(T.outlookSecretAccountIds({ outlookRefreshToken__Work: 'rt-x' }), [], 'an id the rule refuses is not guessed at');
});

test('the six walkers see a second token: moved into the store, held, filled back, resolved, listed, moved through the env file, cleared', async () => {
  const { storage, vault } = store();
  const s = twoSignedIn();
  assert.equal(T.settingsHoldSecrets(s), true);
  assert.deepEqual(T.migrateSecrets(s, vault).moved, FOUR.concat(WORK_FOUR));
  assert.equal(s.outlookRefreshToken__work, '', 'the suffixed field is blanked like the bare one');
  assert.deepEqual(WORK_KEYS.map((k) => storage.getSecret(k)), ['rt-w1', 'at-w1', '2000', 'Work mailbox']);
  assert.deepEqual(DEFAULT_KEYS.map((k) => storage.getSecret(k)), ['rt-d1', 'at-d1', '1000', 'Personal mailbox']);
  assert.equal(s.outlookScopes__work, 'Mail.Read Mail.ReadWrite Calendars.Read', 'the scopes are settings, not a secret: they stay');
  assert.equal(T.settingsHoldSecrets(s), false);
  assert.deepEqual(T.migrateSecrets(s, vault), { changed: false, moved: [] }, 'idempotent');
  assert.ok(!/rt-|at-/.test(JSON.stringify(s)), 'data.json carries no token of either account');
  const w = T.withSecrets(s, vault);
  assert.equal(w.outlookRefreshToken__work, 'rt-w1', 'the resolver fills the suffixed field back in');
  assert.equal(s.outlookRefreshToken__work, '', 'the original stays blank');
  assert.equal(T.dataJsonStore(Object.assign({}, s, { outlookRefreshToken__work: 'rt-w2' })).getSecret(WORK_KEYS[0]), 'rt-w2', 'the data.json store resolves the suffixed id');
  assert.equal(T.dataJsonStore(s).getSecret(`${PREFIX}outlook-refresh-token-nope`), null);
  const slots = T.secretSlots(s);
  assert.deepEqual(slots.map((x) => x.label).slice(5), ['Outlook refresh token (Work)', 'Outlook access token (Work)'], 'two more rows after the five, named by the label');
  assert.deepEqual(slots.map((x) => x.envKey).slice(5), ['OUTLOOK_REFRESH_TOKEN_WORK', 'OUTLOOK_ACCESS_TOKEN_WORK']);
  assert.deepEqual(slots[6].ids, WORK_KEYS.slice(1), 'the expiry and the account name travel with the access token');
  assert.deepEqual(slots[4].ids, DEFAULT_KEYS.slice(1), 'the default rows are what they were');
  assert.deepEqual(T.secretSlots({ outlookRefreshToken__old: 'rt-o1' }).map((x) => x.label).slice(5), ['Outlook refresh token (old)', 'Outlook access token (old)'], 'an unlisted account keeps its rows, named by the id');
  // The settled migrator (the env file's path): its own lines.
  const lines = [];
  const env = { getSecret: (id) => (lines.find((x) => x.id === id) || {}).v || null, setSecret: async (id, v) => { lines.push({ id, v }); return true; }, settle: async () => {} };
  const e = twoSignedIn();
  assert.deepEqual((await T.migrateSecretsSettled(e, new T.SecretVault(null, env))).moved, FOUR.concat(WORK_FOUR));
  assert.deepEqual(lines.map((l) => T.envKeyFor(l.id)).slice(4), ['OUTLOOK_REFRESH_TOKEN_WORK', 'OUTLOOK_ACCESS_TOKEN_WORK', 'OUTLOOK_EXPIRES_AT_WORK', 'OUTLOOK_ACCOUNT_WORK']);
  assert.equal(e.outlookRefreshToken__work, '');
  // Cleared through the sink with the account named: the other account's keys stay.
  T.clearOutlookTokens({ live: s, vault, account: 'work' });
  assert.deepEqual(WORK_KEYS.map((k) => storage.getSecret(k)), ['', '', '', '']);
  assert.equal(storage.getSecret(DEFAULT_KEYS[0]), 'rt-d1', 'the default is untouched');
  T.clearOutlookTokens({ live: s, vault });
  assert.equal(storage.getSecret(DEFAULT_KEYS[0]), '', 'a sink without an account clears the default, as before');
});

test('the view: the default IS the settings object; a further account is the flat shape with its own values; a rotation lands on its own keys', async () => {
  const { storage, vault } = store();
  const s = twoSignedIn();
  T.migrateSecrets(s, vault);
  const w = T.withSecrets(s, vault);
  for (const d of ['default', undefined, TWO[0]]) assert.equal(T.outlookAccountView(w, d), w, 'the same object, not a copy');
  const v = T.outlookAccountView(w, T.outlookAccountById(w, 'work'));
  assert.notEqual(v, w);
  assert.equal(v.outlookClientId, CLIENT, 'the one app registration serves every account');
  assert.deepEqual(T.outlookTokens(v), { refreshToken: 'rt-w1', accessToken: 'at-w1', expiresAt: 2000, account: 'Work mailbox' });
  assert.equal(v.outlookScopes, 'Mail.Read Mail.ReadWrite Calendars.Read');
  assert.equal(T.outlookHasWriteScope(v), true);
  assert.equal(T.outlookHasWriteScope(w), false, 'per account: the default was granted read only');
  assert.equal(T.outlookStatusText(v), 'Signed in as Work mailbox.');
  assert.equal(v._live, s, 'the hidden link points at the live settings, not at the copy');
  assert.equal(v._vault, vault);
  assert.equal(v._account, 'work');
  assert.ok(!Object.keys(v).some((k) => k.startsWith('_')), 'non-enumerable');
  assert.equal(v._shadow, w._shadow, 'the shadow map is the live one');
  assert.deepEqual(T.outlookTokenSink(v), { live: s, vault, view: v, account: 'work' });
  assert.equal(T.outlookTokenSink(w).account, 'default');
  assert.equal(T.outlookSignedIn(T.outlookAccountView(w, 'Work')), false, 'an id the rule refuses reads as not signed in');
  assert.equal(T.outlookSignedIn(T.outlookAccountView({ outlookClientId: CLIENT, outlookAccounts: TWO }, 'work')), false, 'listed but never signed in');
  // A refresh through the work view rotates the work keys and nothing else.
  const x = wire([json(200, { access_token: 'at-w2', refresh_token: 'rt-w2', expires_in: 3600 })]);
  assert.equal(await T.ensureAccessToken(v, { requestUrl: x.requestUrl, now: () => 5000 }, true), 'at-w2');
  assert.equal(form(x.calls[0]).refresh_token, 'rt-w1', 'the work token was sent');
  assert.match(form(x.calls[0]).scope, /Mail\.ReadWrite/, 'with the work scopes');
  assert.deepEqual(WORK_KEYS.slice(0, 3).map((k) => storage.getSecret(k)), ['rt-w2', 'at-w2', String(5000 + 3600000)]);
  assert.deepEqual(DEFAULT_KEYS.slice(0, 2).map((k) => storage.getSecret(k)), ['rt-d1', 'at-d1'], 'the default keys did not move');
  assert.equal(v.outlookRefreshToken, 'rt-w2', 'the view the run holds reads the rotated token');
  assert.equal(w.outlookRefreshToken, 'rt-d1');
  assert.equal(s.outlookRefreshToken__work, '', 'the live settings stay blank in store mode');
  // Without a store the rotation lands on the suffixed field of the live settings and asks for a save.
  const plain = twoSignedIn();
  let persisted = 0;
  Object.defineProperty(plain, '_persist', { value: () => { persisted += 1; }, enumerable: false });
  const y = wire([json(200, { access_token: 'at-w3', refresh_token: 'rt-w3', expires_in: 60 })]);
  await T.ensureAccessToken(T.outlookAccountView(T.withSecrets(plain, new T.SecretVault(null)), 'work'), { requestUrl: y.requestUrl, now: () => 7 }, true);
  assert.deepEqual([plain.outlookRefreshToken__work, plain.outlookAccessToken__work, plain.outlookRefreshToken, persisted], ['rt-w3', 'at-w3', 'rt-d1', 1]);
});

// A stand-in for the sign-in dialog: what was written into it, whether it closed.
const fakeModal = (pending) => ({ pending, statuses: [], closed: false, setStatus(t, failed) { this.statuses.push([t, !!failed]); }, close() { this.closed = true; } });

test('one sign-in in flight, found by its state: starting Work forgets Personal; a reply from the older tab exchanges nothing and leaves the newer dialog alone', async () => {
  const p = headless(twoSignedIn(), new T.SecretVault(null));
  const t0 = Date.now() - 1000; // the reply sweeps by Date.now()
  const a = p.outlookRememberPending({ state: 'st-a', verifier: 'v-a', clientId: CLIENT, tenant: 'common', scopes: 'a', accountId: 'default' }, t0);
  // What outlookSignIn does for the second sign-in: forget every other, remember this one, open its own dialog.
  p.outlookClearPending();
  const b = p.outlookRememberPending({ state: 'st-b', verifier: 'v-b', clientId: CLIENT, tenant: 'common', scopes: 'b', accountId: 'work' }, t0 + 1000);
  p._outlookModal = fakeModal(b);
  assert.deepEqual([...p.outlookPendingMap().keys()], ['st-b'], 'starting Work dropped Personal');
  const finished = [];
  p.outlookFinishSignIn = async (tokens, pending) => { finished.push({ tokens, pending }); };
  // The older Personal tab finishes: no exchange, nothing stored, and the Work dialog untouched.
  const quiet = wire([]);
  await p.outlookAuthCallback({ code: 'code-a', state: 'st-a' }, { requestUrl: quiet.requestUrl });
  assert.equal(quiet.calls.length, 0, 'a reply for a forgotten sign-in makes no exchange');
  assert.equal(finished.length, 0);
  assert.deepEqual(p._outlookModal.statuses, [], 'the Work dialog heard nothing');
  assert.equal(p._outlookModal.closed, false);
  for (const params of [{ code: 'code-x', state: 'st-x' }, { code: 'code-x' }, { error: 'access_denied', error_description: 'AADSTS65004: User declined to consent.', state: 'st-b' }]) {
    await p.outlookAuthCallback(params, { requestUrl: quiet.requestUrl });
  }
  assert.equal(quiet.calls.length, 0, 'an unknown state, no state, or a declined reply: no exchange');
  assert.deepEqual([...p.outlookPendingMap().keys()], ['st-b'], 'and nothing consumed: Work waits for the next try');
  assert.deepEqual(p._outlookModal.statuses.map((s) => s[1]), [true], 'only the declined reply for Work itself reached the Work dialog');
  // Then Work's own reply.
  const x = wire([json(200, { access_token: 'at-w1', refresh_token: 'rt-w1', expires_in: 3600, scope: 'b' })]);
  await p.outlookAuthCallback({ action: 'icor-for-life-planner/auth', code: 'code-b', state: 'st-b' }, { requestUrl: x.requestUrl });
  assert.equal(finished.length, 1);
  assert.equal(finished[0].pending, b);
  assert.equal(form(x.calls[0]).code_verifier, 'v-b', 'exchanged with Work\'s own verifier');
  assert.equal(p.outlookPendingMap().size, 0);
  assert.equal(a.state, 'st-a', 'the forgotten entry is just an object now');
  // The TTL, and a sign-out forgets its own account's sign-ins only.
  assert.equal(T.OUTLOOK_PENDING_TTL_MS, 15 * 60000);
  p.outlookRememberPending({ state: 'st-old', accountId: 'work' }, 0);
  p.outlookRememberPending({ state: 'st-new', accountId: 'work' }, T.OUTLOOK_PENDING_TTL_MS + 1);
  p.outlookClearPending('default');
  assert.deepEqual([...p.outlookPendingMap().keys()], ['st-new']);
  p.outlookClearPending('work');
  assert.equal(p.outlookPendingMap().size, 0);
  // A sign-in for an account the list does not carry starts nothing.
  const before = p._outlookModal;
  await p.outlookSignIn({ accountId: 'nope' });
  await p.outlookSignIn({ accountId: 'Work' });
  assert.equal(p.outlookPendingMap().size, 0);
  assert.equal(p._outlookModal, before, 'and opens no dialog');
});

test('a further account\'s sign-in stores under its keys and leaves the default\'s; sign-out of one account leaves the other', async () => {
  const { storage, vault } = store();
  const s = Object.assign(twoSignedIn(), { outlookRefreshToken__work: '', outlookAccessToken__work: '', outlookExpiresAt__work: '', outlookAccount__work: '', outlookScopes__work: '' });
  T.migrateSecrets(s, vault);
  const p = headless(s, vault);
  let done = 0;
  const me = wire([json(200, { userPrincipalName: 'work-mailbox', displayName: 'Work' })]);
  const pending = { state: 'st-w', accountId: 'work', scopes: 'a', onDone: () => { done += 1; } };
  const other = fakeModal({ state: 'st-other', accountId: 'default' });
  p._outlookModal = other;
  await p.outlookFinishSignIn({ accessToken: 'at-w1', refreshToken: 'rt-w1', expiresIn: 3600, scope: 'Mail.Read Calendars.Read' }, pending, { requestUrl: me.requestUrl, now: () => 9000 });
  assert.equal(other.closed, false, 'another sign-in\'s dialog is not taken down');
  assert.equal(p._outlookModal, other);
  p._outlookModal = fakeModal(pending);
  p.outlookCloseModalFor(pending);
  assert.equal(p._outlookModal, null, 'its own dialog closes');
  assert.match(me.calls[0].url, /\/me\?/);
  assert.equal(me.calls[0].headers.Authorization, 'Bearer at-w1', 'the /me call carries the new account\'s token');
  const stored = WORK_KEYS.map((k) => storage.getSecret(k));
  assert.deepEqual([stored[0], stored[1], stored[3]], ['rt-w1', 'at-w1', 'work-mailbox']);
  assert.ok(Number(stored[2]) > Date.now() + 3500000, 'the expiry is stamped off the real clock, as the default\'s is');
  assert.deepEqual(DEFAULT_KEYS.map((k) => storage.getSecret(k)), ['rt-d1', 'at-d1', '1000', 'Personal mailbox'], 'the default keys are what they were');
  assert.equal(s.outlookScopes__work, 'Mail.Read Calendars.Read', 'the granted scopes, per account');
  assert.equal(s.outlookScopes, 'Mail.Read Calendars.Read', 'the default\'s scopes untouched');
  assert.equal(done, 1);
  assert.equal(p.saved.length, 1);
  assert.ok(!/rt-|at-/.test(p.saved[0]), 'no token in what reaches disk');
  assert.equal(p.synced, 1, 'a further account syncs on its own run once signed in (part 3)');
  assert.deepEqual(s.calendars.map((f) => f.id), ['outlook-graph', 'outlook-graph-work'], 'and its own calendar feed, never under the default\'s id (part 3)');
  assert.equal(T.outlookSignedIn(T.outlookAccountView(T.withSecrets(s, vault), 'work')), true);
  // Sign out of work: its four keys and its scopes go, the default stays.
  await p.outlookSignOut('work');
  assert.deepEqual(WORK_KEYS.map((k) => storage.getSecret(k)), ['', '', '', '']);
  assert.equal(storage.getSecret(DEFAULT_KEYS[0]), 'rt-d1');
  assert.ok(!Object.keys(s).some((k) => k.endsWith('__work')), 'the suffixed fields leave data.json');
  assert.equal(s.outlookScopes, 'Mail.Read Calendars.Read');
  assert.equal(T.outlookSignedIn(T.withSecrets(s, vault)), true, 'the default is still signed in');
  assert.equal(p.recomputed, 1, 'its calendar events leave the board on sign-out, as the default\'s do (part 3)');
  // Sign out of the default: the literal legacy path; the other account's key stays.
  storage.setSecret(WORK_KEYS[0], 'rt-w9');
  await p.outlookSignOut();
  assert.deepEqual([storage.getSecret(DEFAULT_KEYS[0]), storage.getSecret(DEFAULT_KEYS[3]), s.outlookScopes, p.recomputed], ['', '', '', 2]);
  assert.equal(storage.getSecret(WORK_KEYS[0]), 'rt-w9', 'the other account\'s key stays');
  await p.outlookSignOut('Work');
  assert.equal(storage.getSecret(WORK_KEYS[0]), 'rt-w9', 'an id the rule refuses signs nobody out');
});

test('source scan: the walkers on secretFieldNames(), the pins the sign-in keeps, the settings rows', () => {
  const c = code();
  const r = raw();
  assert.equal((c.match(/\bSECRET_FIELD_NAMES\b/g) || []).length, 2, 'declared once, read once (by secretFieldNames); no walker reads the constant');
  for (const fn of ['function migrateSecrets(settings, vault)', 'function settingsHoldSecrets(settings)', 'async function migrateSecretsSettled(settings, vault)', 'function withSecrets(settings, vault)', 'function dataJsonStore(settings)']) {
    const at = c.indexOf(fn);
    assert.ok(at > 0, fn);
    assert.match(c.slice(at, c.indexOf('\n}\n', at)), /secretFieldNames\(s\)/, `${fn} walks secretFieldNames`);
  }
  assert.match(c, /function secretSlots\(settings\) \{[\s\S]{0,900}outlookSecretAccountIds\(settings\)/, 'the key list adds the account rows');
  assert.match(c, /writeSecret\(live, k\.vault, outlookAccountField\(acct, field\), value\)/, 'every token write names the account');
  assert.match(c, /writeSecret\(live, k\.vault, outlookAccountField\(acct, f\), ''\)/, 'and every clear');
  // The pins: one handler, four resolved views in the class, the literal default calls.
  // One protocol handler, registered as an arrow that passes `params` and
  // nothing else: production always gets the real requestUrl, and nothing
  // off an obsidian:// URL can reach the trailing `deps`.
  assert.equal((c.match(/registerObsidianProtocolHandler\(/g) || []).length, 1);
  assert.match(c, /this\.registerObsidianProtocolHandler\(OUTLOOK_PROTOCOL_ACTION, \(params\) => this\.outlookAuthCallback\(params\)\);/);
  assert.doesNotMatch(c, /outlookAuthCallback\.bind\(|outlookAuthCallback\(\.\.\./);
  assert.equal((c.match(/const s = this\.withSecrets\(\);/g) || []).length, 4);
  assert.match(c, /clearOutlookTokens\(\{ live: this\.settings, vault: this\.secrets \}\)/, 'the default sign-out is the literal call');
  assert.match(c, /clearOutlookTokens\(\{ live: this\.settings, vault: this\.secrets, account: id \}\)/, 'a further account takes the widened one');
  assert.equal((c.match(/ensureGraphCalendarFeed\(/g) || []).length, 3, 'the definition, the literal default call, and the widened call for a further account (part 3)');
  assert.match(c, /ensureGraphCalendarFeed\(this\.settings\);/, 'the default\'s call is the literal one');
  assert.doesNotMatch(c, /this\._outlookPending = (\{|null)/, 'no single latest sign-in any more');
  assert.doesNotMatch(c, /this\.(plugin\.)?settings\.outlook(RefreshToken|AccessToken|ExpiresAt|Account)(__|\b)/, 'no class reads a token off the settings, suffixed or not');
  assert.match(c, /const pending = modal \? modal\.pending : null;/, 'the device code finishes the account the modal was opened for');
  const defaults = r.slice(r.indexOf('const DEFAULT_SETTINGS = {'), r.indexOf('\n};', r.indexOf('const DEFAULT_SETTINGS = {')));
  assert.doesNotMatch(defaults, /outlookAccounts|__/, 'no default under the list and no suffixed field is laid down');
  // The settings rows: Setting rows, no raw heading tag, no inline style,
  // plain ASCII; refreshed with the default row from one resolved copy,
  // never by a re-render per keystroke.
  const tab = r.slice(r.indexOf('class IcorPlannerSettingTab'));
  const rows = tab.slice(tab.indexOf('for (const account of accounts.slice(1))'), tab.indexOf("setName('Manage or revoke access')"));
  assert.ok(rows.length > 0);
  assert.match(rows, /new Setting\(containerEl\)\.setName\(`Microsoft account: \$\{account\.label\}`\)/);
  assert.match(rows, /accountRefreshers\.push\(\(r\) => \{\n\s*const view = outlookAccountView\(r, account\);/);
  assert.match(tab, /if \(signOutBtn\) signOutBtn\.setDisabled\(!signed\);\n\s*for \(const refresh of accountRefreshers\) refresh\(r\);/, 'the default row\'s refresh drives the further rows');
  assert.doesNotMatch(rows, /createEl\('h[1-6]'|\.style\.|style=|setHeading|withSecrets\(\)/, 'no heading tag, no inline style, no second read of the store');
  assert.ok([...rows].every((ch) => ch.charCodeAt(0) < 128), 'plain ASCII, so no dash of either length');
  for (const text of rows.match(/setButtonText\('([^']+)'\)/g) || []) assert.match(text, /'(Sign in|Sign in again|Sign out|Add)'/, 'sentence case (Add is part 4b\'s row, after the account rows)');
  assert.ok(tab.indexOf('/* ---- Outlook (2026-09-06) ---- */') > tab.indexOf("setName('Starred email (IMAP)').setHeading()"), 'the marker stays after the IMAP heading');
});
