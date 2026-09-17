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
const freezeErrors: string[] = [];

function makeRange(r: number, c: number, nr: number, nc: number, merges: number[][] = []) {
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
    merge: () => { check(nr >= 1 && nc >= 1 && nr * nc > 1, `병합 범위 ${r},${c},${nr},${nc}`); merges.push([r, c, r + nr - 1, c + nc - 1]); calls.push('merge'); return rng; },
  };
  check(r >= 1 && c >= 1 && nr >= 1 && nc >= 1, `getRange 범위 ${r},${c},${nr},${nc}`);
  return rng;
}
function makeSheet() {
  let cols = 26, rows = 1000;
  const merges: number[][] = [];
  const freezeGuard = (kind: string, n: number) => {
    const bad = merges.some(([r1, c1, r2, c2]) => (kind === 'col' ? c1 <= n && c2 > n : r1 <= n && r2 > n));
    if (bad) { freezeErrors.push(`${kind}:${n}`); throw new Error('병합된 셀의 일부만 포함된 열을 고정할 수 없습니다.'); }
    calls.push(`freeze-${kind}:${n}`);
  };
  const sh: any = {
    setName: (n: string) => { calls.push(`sheet:${n}`); return sh; },
    getMaxColumns: () => cols, getMaxRows: () => rows,
    insertColumnsAfter: (_a: number, n: number) => { cols += n; }, insertRowsAfter: (_a: number, n: number) => { rows += n; },
    getRange: (r: number, c: number, nr: number, nc: number) => { check(c + nc - 1 <= cols && r + nr - 1 <= rows, '시트 크기 초과'); return makeRange(r, c, nr, nc, merges); },
    setColumnWidth: () => sh, setRowHeight: () => sh,
    setFrozenRows: (n: number) => { freezeGuard('row', n); return sh; }, setFrozenColumns: (n: number) => { freezeGuard('col', n); return sh; },
  };
  return sh;
}
const SpreadsheetApp = {
  BorderStyle: { SOLID: 'SOLID' },
  create: (title: string) => { calls.push(`create:${title}`); const first = makeSheet(); return { getSheets: () => [first], insertSheet: () => makeSheet(), getUrl: () => 'https://docs.google.com/spreadsheets/d/X/edit', getId: () => 'X' }; },
};
const DriveApp = {
  Access: { ANYONE_WITH_LINK: 'A' }, Permission: { VIEW: 'V' },
  getFileById: () => ({ moveTo: () => calls.push('moveTo'), setSharing: () => calls.push('setSharing'), setTrashed: () => calls.push('trashed') }),
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

const { crossesFreeze } = await import('../src/daySheet');
book.sheets.forEach((sp) => {
  check(!(sp.frozenCols && crossesFreeze(sp.merges, 'col', sp.frozenCols)), `${sp.name}: 병합을 가로지르는 열 고정`);
  check(!(sp.frozenRows && crossesFreeze(sp.merges, 'row', sp.frozenRows)), `${sp.name}: 병합을 가로지르는 행 고정`);
});
const ping = JSON.parse(gs.doGet().body);
check(ping.ok && ping.app === 'proctor-sheet', 'doGet 응답');
const bad = JSON.parse(gs.doPost({ postData: { contents: JSON.stringify({ token: 'wrong', book }) } }).body);
check(!bad.ok && /토큰/.test(bad.error), '토큰 불일치 거부');
const ok = JSON.parse(gs.doPost({ postData: { contents: JSON.stringify({ token: 'tok123', folderId: 'F', shareLink: true, book }) } }).body);
check(ok.ok === true && /spreadsheets/.test(ok.url), `doPost 응답 ${JSON.stringify(ok)}`);
check(calls.includes('moveTo') && calls.includes('setSharing'), '폴더 이동·링크 공유 호출');
check(freezeErrors.length === 0, `구글이 거부할 고정: ${freezeErrors}`);
check(!calls.includes('trashed'), '정상 생성인데 휴지통 이동이 호출됨');
console.log('고정:', calls.filter((c) => c.startsWith('freeze')).join(', ') || '없음');
check(calls.filter((c) => c.startsWith('sheet:')).length === book.sheets.length, '시트 수');
console.log('시트:', calls.filter((c) => c.startsWith('sheet:')).join(', '), '/ 병합', calls.filter((c) => c === 'merge').length, '회 / 응답', ok.ok);

// 선생님이 이미 배포한 이전 스크립트(고정 실패를 그대로 오류로 내던 버전)를 재현
const legacyCode = code
  .replace('try { if (spec.frozenRows) sh.setFrozenRows(spec.frozenRows); } catch (ignore) {}', 'if (spec.frozenRows) sh.setFrozenRows(spec.frozenRows);')
  .replace('try { if (spec.frozenCols) sh.setFrozenColumns(spec.frozenCols); } catch (ignore) {}', 'if (spec.frozenCols) sh.setFrozenColumns(spec.frozenCols);');
check(legacyCode !== code && !legacyCode.includes('sh.setFrozenColumns(spec.frozenCols); } catch'), '이전 스크립트 재현 실패');
const legacy = new Function('SpreadsheetApp', 'DriveApp', 'ContentService', `${legacyCode}\nreturn { doGet, doPost };`)(SpreadsheetApp, DriveApp, ContentService);

// (1) 수정 전 구조(시감표 첫 열 고정 + 제목 병합)면 선생님이 받은 오류가 그대로 나야 한다
const oldBook = JSON.parse(JSON.stringify(book));
oldBook.sheets[0].frozenCols = 1;
const oldRes = JSON.parse(legacy.doPost({ postData: { contents: JSON.stringify({ token: 'tok123', book: oldBook }) } }).body);
check(!oldRes.ok && /병합된 셀의 일부만 포함된 열을 고정할 수 없습니다/.test(oldRes.error), `수정 전 오류 재현 실패: ${JSON.stringify(oldRes)}`);
console.log('수정 전 구조 + 이전 스크립트:', oldRes.ok ? '성공(재현 안 됨)' : `오류 재현 → ${oldRes.error}`);

// (2) 수정 후 구조는 이전 스크립트로도 모든 날짜가 만들어져야 한다
freezeErrors.length = 0;
const days = [...new Set(slots.map((s) => s.dayIdx))];
days.forEach((d) => {
  const r = JSON.parse(legacy.doPost({ postData: { contents: JSON.stringify({ token: 'tok123', book: buildDayBook(p, slots, d) }) } }).body);
  check(r.ok === true, `이전 스크립트로 ${d + 1}일차 생성 실패: ${r.error}`);
});
check(freezeErrors.length === 0, `이전 스크립트에서 고정 오류: ${freezeErrors}`);
console.log(`수정 후 구조 + 이전 스크립트: ${days.length}일 모두 생성 ${freezeErrors.length ? '실패' : '성공'}`);

console.log(failed ? `\n${failed} FAILED` : '\nALL CHECKS PASSED');
process.exit(failed ? 1 : 0);
