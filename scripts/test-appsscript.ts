// 구글 앱스 스크립트 코드를 흉내 객체로 실행해 호출 이름·배열 크기를 검증
(globalThis as any).document = { createElement: () => ({ click() {}, remove() {} }), body: { appendChild() {} } };
const { appsScriptCode } = await import('../src/appsScript');
const { buildDayBook } = await import('../src/daySheet');
const { demoProject } = await import('../src/demo');
const { applyGrid, autoFillTargets, buildEngineInput, getSlots } = await import('../src/model');
const { runBest } = await import('../src/engine');

let failed = 0;
const check = (cond: boolean, msg: string) => { if (!cond) { failed++; console.log('  FAIL:', msg); } };
const calls: string[] = [];

function makeRange(r: number, c: number, nr: number, nc: number) {
  const shape = (name: string) => (arr: unknown[][]) => {
    check(Array.isArray(arr) && arr.length === nr && arr.every((row) => Array.isArray(row) && row.length === nc), `${name} 배열 크기 ${arr?.length}x${(arr?.[0] as any)?.length} != ${nr}x${nc}`);
    calls.push(name);
    return rng;
  };
  const rng: any = {
    setNumberFormat: () => rng, setValues: shape('setValues'), setBackgrounds: shape('setBackgrounds'), setFontWeights: shape('setFontWeights'),
    setFontColors: shape('setFontColors'), setHorizontalAlignments: shape('setHorizontalAlignments'), setWraps: shape('setWraps'), setFontSizes: shape('setFontSizes'),
    setVerticalAlignment: () => rng,
    setBorder: (...a: unknown[]) => { check(a.length === 8, 'setBorder 인자 수'); return rng; },
    merge: () => { check(nr >= 1 && nc >= 1 && nr * nc > 1, `병합 범위 ${r},${c},${nr},${nc}`); calls.push('merge'); return rng; },
  };
  check(r >= 1 && c >= 1 && nr >= 1 && nc >= 1, `getRange 범위 ${r},${c},${nr},${nc}`);
  return rng;
}
function makeSheet() {
  let cols = 26, rows = 1000;
  const sh: any = {
    setName: (n: string) => { calls.push(`sheet:${n}`); return sh; },
    getMaxColumns: () => cols, getMaxRows: () => rows,
    insertColumnsAfter: (_a: number, n: number) => { cols += n; }, insertRowsAfter: (_a: number, n: number) => { rows += n; },
    getRange: (r: number, c: number, nr: number, nc: number) => { check(c + nc - 1 <= cols && r + nr - 1 <= rows, '시트 크기 초과'); return makeRange(r, c, nr, nc); },
    setColumnWidth: () => sh, setRowHeight: () => sh, setFrozenRows: () => sh, setFrozenColumns: () => sh,
  };
  return sh;
}
const SpreadsheetApp = {
  BorderStyle: { SOLID: 'SOLID' },
  create: (title: string) => { calls.push(`create:${title}`); const first = makeSheet(); return { getSheets: () => [first], insertSheet: () => makeSheet(), getUrl: () => 'https://docs.google.com/spreadsheets/d/X/edit', getId: () => 'X' }; },
};
const DriveApp = {
  Access: { ANYONE_WITH_LINK: 'A' }, Permission: { VIEW: 'V' },
  getFileById: () => ({ moveTo: () => calls.push('moveTo'), setSharing: () => calls.push('setSharing') }),
  getFolderById: () => ({}),
};
const ContentService = { MimeType: { JSON: 'json' }, createTextOutput: (s: string) => ({ setMimeType: () => ({ body: s }) }) };

const code = appsScriptCode('tok123');
const gs = new Function('SpreadsheetApp', 'DriveApp', 'ContentService', `${code}\nreturn { doGet, doPost };`)(SpreadsheetApp, DriveApp, ContentService);

let p = demoProject();
const slots = getSlots(p);
p = autoFillTargets(p, slots).project;
const res = runBest(buildEngineInput(p, slots), 2, 4);
if (res.status !== 'ok') throw new Error(res.status);
p = applyGrid(p, slots, res.grid);
p.exams = { [`${slots[0].key}|2학년`]: '화법과 언어 (40분/공통) (8:20~9:00) 교사01, 교사02' };
const book = buildDayBook(p, slots, 0);

const ping = JSON.parse(gs.doGet().body);
check(ping.ok && ping.app === 'proctor-sheet', 'doGet 응답');
const bad = JSON.parse(gs.doPost({ postData: { contents: JSON.stringify({ token: 'wrong', book }) } }).body);
check(!bad.ok && /토큰/.test(bad.error), '토큰 불일치 거부');
const ok = JSON.parse(gs.doPost({ postData: { contents: JSON.stringify({ token: 'tok123', folderId: 'F', shareLink: true, book }) } }).body);
check(ok.ok === true && /spreadsheets/.test(ok.url), `doPost 응답 ${JSON.stringify(ok)}`);
check(calls.includes('moveTo') && calls.includes('setSharing'), '폴더 이동·링크 공유 호출');
check(calls.filter((c) => c.startsWith('sheet:')).length === book.sheets.length, '시트 수');
console.log('시트:', calls.filter((c) => c.startsWith('sheet:')).join(', '), '/ 병합', calls.filter((c) => c === 'merge').length, '회 / 응답', ok.ok);
console.log(failed ? `\n${failed} FAILED` : '\nALL CHECKS PASSED');
process.exit(failed ? 1 : 0);
