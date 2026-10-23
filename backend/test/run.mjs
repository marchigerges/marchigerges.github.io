// Runs backend/Code.gs against an in-memory copy of the Sheet layout.
// Usage: node backend/test/run.mjs
import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';

const src = fs.readFileSync(new URL('../Code.gs', import.meta.url), 'utf8');

function makeSheet(name, rows) {
  const data = rows.map(r => r.slice());
  const width = () => Math.max(0, ...data.map(r => r.length));
  const cell = (r, c) => (data[r] && data[r][c] !== undefined ? data[r][c] : '');
  const put = (r, c, v) => { while (data.length <= r) data.push([]); while (data[r].length <= c) data[r].push(''); data[r][c] = v; };
  const range = (row, col, nr = 1, nc = 1) => ({
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => cell(row - 1 + i, col - 1 + j))),
    setValues: v => v.forEach((r, i) => r.forEach((x, j) => put(row - 1 + i, col - 1 + j, x))),
    setValue: v => put(row - 1, col - 1, v),
  });
  const api = {
    name, data,
    getRange: range,
    getDataRange: () => range(1, 1, data.length, width()),
    getLastRow: () => { for (let i = data.length - 1; i >= 0; i--) if (data[i].some(x => x !== '')) return i + 1; return 0; },
    getLastColumn: width,
    appendRow: r => { data[api.getLastRow()] = r.slice(); },
    deleteRow: r => data.splice(r - 1, 1),
    setFrozenRows: () => {},
  };
  return api;
}

const sheets = {};
function add(name, rows) { sheets[name] = makeSheet(name, rows); }

add('Invites', [
  ['Code', 'Side', 'Group', 'Guest 1', 'Guest 2', 'Guest 3', 'Guest 4', 'Guest 5', 'Guest 6', 'Personal link', 'Check', 'Replied', 'Site address', 'https://x'],
  ['E9mqXIKdiW', 'Salwa', 'Salwa Family', 'Joseph El Chami', 'Noura El Chami'],
  ['qje1H2qBb4', 'Salwa', 'Extended Family', 'Abou Anis', "Abou Anis's wife"],
  ['h3wetsfTbX', 'Salwa', 'Jessica Jebbo Family', 'Jessica Jebbo'],
]);
add('Guest List', [
  ['Salwa Side', '', '', '', '', '', 'Marchi Side'],
  ['Group', 'Name', 'RSVP', 'Allergies', 'Notes', '', 'Group', 'Name', 'RSVP', 'Allergies', 'Notes'],
  ['Salwa Family', 'Joseph El Chami', '', '', '', '', 'Marchi Family', 'Mariam Bekhtan'],
  ['Test', 'Alex Testguest', '', '', '', '', 'Test', 'Jordan Testguest'],
  ['', '', '', '', '', '', 'Test', 'Sam Plusone'],
]);

const props = {};
let now = new Date('2026-11-01T10:00:00+11:00');
const RealDate = Date;
class FakeDate extends RealDate { constructor(...a) { super(...(a.length ? a : [now])); } static now() { return now.getTime(); } }

const ctx = {
  Date: FakeDate, JSON, String, Array, Object, console,
  SpreadsheetApp: { getActive: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => (add(n, []), sheets[n]) }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: { formatDate: d => d.toISOString().slice(0, 16).replace('T', ' ') },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType: () => ({ body: s }) }) },
};
vm.createContext(ctx);
vm.runInContext(src, ctx);

const get = q => JSON.parse(ctx.doGet({ parameter: q }).body);
const post = b => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(b) } }).body);
const log = (label, v) => console.log(`\n# ${label}\n` + JSON.stringify(v));

// Setup, as the couple would run it from the editor.
props.ACCESS_CODE = 'Riviera';
ctx.addTestGuests();
const responses = sheets.Responses.data;
log('Responses tab after addTestGuests (slugs)', responses.map(r => r.slice(0, 2)));
const slugOf = code => responses.find(r => r[0] === code)[1];
const alex = slugOf('TEST000001'), jordan = slugOf('TEST000002');
assert.equal(alex, 'alex-testguest');
assert.equal(jordan, 'jordan-sam');
assert.equal(slugOf('qje1H2qBb4'), 'abou-abou');

// Gate
let r = get({ slug: 'nobody' }); log('unknown slug', r); assert.equal(r.exists, false);
r = get({ slug: '' }); log('bare domain (no slug)', r); assert.equal(r.exists, false);
r = get({ slug: 'Jordan-Sam' }); log('valid slug, any case', r); assert.equal(r.exists, true);
r = get({ slug: jordan, code: 'wrong' }); log('wrong code', r); assert.equal(r.error, 'denied');
r = get({ slug: 'nobody', code: 'riviera' }); log('right code, unknown slug', r); assert.equal(r.error, 'denied');
r = get({ slug: jordan, code: '  rIVIERA ' }); log('right code, case-insensitive', r);
assert.deepEqual(r.invite.guests, ['Jordan Testguest', 'Sam Plusone']); assert.equal(r.invite.rsvp, null); assert.equal(r.invite.open, true);

// RSVP validation
r = post({ slug: jordan, code: 'riviera', guests: [{ name: 'Jordan Testguest', attending: true }] });
log('POST missing plus-one', r); assert.equal(r.error, 'invalid');
r = post({ slug: jordan, code: 'riviera', guests: [{ name: 'Jordan Testguest', attending: true }, { name: 'Someone Else', attending: true }] });
log('POST renamed plus-one', r); assert.equal(r.error, 'invalid');
r = post({ slug: jordan, code: 'nope', guests: [] }); log('POST wrong code', r); assert.equal(r.error, 'denied');

// Jordan accepts with dietary, Sam declines
r = post({ slug: jordan, code: 'riviera', notes: 'Looking forward to it', guests: [
  { name: 'Jordan Testguest', attending: true, dietary: 'Vegetarian' },
  { name: 'Sam Plusone', attending: false, dietary: 'Vegan' },
] });
log('POST Jordan yes (veg) + Sam no', r);
assert.equal(r.ok, true); assert.equal(r.invite.rsvp.guests[1].dietary, 'None');

r = post({ slug: jordan, code: 'riviera', guests: [{ name: 'Jordan Testguest', attending: false }, { name: 'Sam Plusone', attending: false }] });
log('POST again (locked)', r); assert.equal(r.error, 'already_replied'); assert.equal(r.invite.rsvp.guests[0].attending, true);

r = get({ slug: jordan, code: 'riviera' }); log('GET after reply shows locked RSVP', r.invite.rsvp); assert.ok(r.invite.rsvp);

// Alex: bogus dietary is coerced, then RSVP closes
r = post({ slug: alex, code: 'riviera', guests: [{ name: 'Alex Testguest', attending: true, dietary: '<script>' }] });
log('POST Alex yes, bogus dietary', r.invite.rsvp); assert.equal(r.invite.rsvp.guests[0].dietary, 'None');

now = new Date('2027-01-01T00:00:01+11:00');
r = post({ slug: slugOf('h3wetsfTbX'), code: 'riviera', guests: [{ name: 'Jessica Jebbo', attending: true }] });
log('POST after 31 Dec 2026', r); assert.equal(r.error, 'closed');
r = get({ slug: slugOf('h3wetsfTbX'), code: 'riviera' }); assert.equal(r.invite.open, false);

// Sheet side effects
const respRow = c => responses.find(x => x[0] === c);
log('Responses row, Jordan', respRow('TEST000002').slice(0, 8));
assert.equal(respRow('TEST000002')[3], 'Partial');
const invRow = c => sheets.Invites.data.find(x => x[0] === c);
log('Invites "Replied", Jordan', invRow('TEST000002')[11]);
assert.match(invRow('TEST000002')[11], /^Partial /);
const gl = sheets['Guest List'].data;
log('Guest List rows', gl.slice(3));
assert.equal(gl[3][2], 'Yes'); assert.equal(gl[3][8], 'Yes'); assert.equal(gl[3][9], 'Vegetarian'); assert.equal(gl[4][8], 'No');

ctx.removeTestGuests();
assert.ok(!sheets.Invites.data.some(x => /^TEST/.test(x[0])));
assert.ok(!responses.some(x => /^TEST/.test(x[0])));
console.log('\nremoveTestGuests: test rows gone\n\nALL PASSED');
