/**
 * 구글 앱스 스크립트 웹 앱을 통해 사용자 드라이브에 구글 시트를 만든다.
 * 앱은 브라우저에서만 돌아가므로, 사용자가 자기 계정으로 아래 스크립트를 한 번 배포해 두고
 * 그 웹 앱 주소로 시감표 구조(JSON)를 보내면 스크립트가 시트를 만들어 주소를 돌려준다.
 */
import type { SheetBook } from './daySheet';

export interface SheetSettings {
  url: string;
  token: string;
  folderId: string;
  shareLink: boolean;
}

const KEY = 'proctor-app:google-sheet';
const EMPTY: SheetSettings = { url: '', token: '', folderId: '', shareLink: false };

export function loadSheetSettings(): SheetSettings {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) return { ...EMPTY, ...JSON.parse(raw) };
  } catch {
    /* 저장된 설정 없음 */
  }
  return { ...EMPTY };
}

export function saveSheetSettings(s: SheetSettings) {
  try {
    localStorage.setItem(KEY, JSON.stringify(s));
  } catch {
    /* 저장 불가 환경 */
  }
}

export const isWebAppUrl = (u: string) => /^https:\/\/script\.google\.com\/(a\/[^/]+\/)?macros\/s\/[\w-]+\/exec\/?$/.test(u.trim());

/** 폴더 주소를 붙여넣어도 ID만 뽑아 쓴다 */
export const folderIdOf = (v: string) => {
  const m = /folders\/([\w-]+)/.exec(v);
  return (m ? m[1] : v).trim();
};

export function randomToken() {
  const a = new Uint8Array(12);
  crypto.getRandomValues(a);
  return [...a].map((b) => b.toString(16).padStart(2, '0')).join('');
}

async function readJson(res: Response) {
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    throw new Error('웹 앱 응답을 읽지 못했습니다. 배포할 때 액세스 권한을 "모든 사용자"로 했는지, URL이 /exec로 끝나는지 확인하세요.');
  }
}

export async function testConnection(url: string) {
  let res: Response;
  try {
    res = await fetch(url.trim());
  } catch {
    throw new Error('웹 앱에 연결하지 못했습니다. 인터넷 연결과 웹 앱 URL을 확인하세요.');
  }
  const data = await readJson(res);
  if (!data.ok || data.app !== 'proctor-sheet') throw new Error('이 앱용 스크립트가 아닙니다. 설정 창의 코드를 다시 붙여넣고 새 버전으로 배포하세요.');
  return data as { ok: true; app: string; version: number };
}

export async function createGoogleSheet(s: SheetSettings, book: SheetBook): Promise<{ url: string; id: string }> {
  let res: Response;
  try {
    // Content-Type을 지정하지 않아 text/plain 단순 요청으로 보낸다(앱스 스크립트는 사전 요청을 받지 않음)
    res = await fetch(s.url.trim(), {
      method: 'POST',
      body: JSON.stringify({ token: s.token, folderId: folderIdOf(s.folderId) || undefined, shareLink: s.shareLink, book }),
    });
  } catch {
    throw new Error('웹 앱에 연결하지 못했습니다. 인터넷 연결과 웹 앱 URL을 확인하세요.');
  }
  const data = await readJson(res);
  if (!data.ok) throw new Error(data.error || '구글 시트를 만들지 못했습니다.');
  return data;
}

/** 사용자가 script.google.com에 붙여넣을 코드 (연결 토큰 포함) */
export function appsScriptCode(token: string): string {
  return `/**
 * 정기고사 시감표 — 구글 시트 만들기 (Google Apps Script)
 *
 * 배포: 오른쪽 위 [배포] > [새 배포] > 유형 [웹 앱]
 *       실행 사용자: 나 / 액세스 권한이 있는 사용자: 모든 사용자
 * 코드를 고치거나 토큰을 바꾸면 [배포 관리]에서 새 버전으로 다시 배포하세요.
 */
var SHARED_TOKEN = ${JSON.stringify(token)};

function doGet() {
  return json_({ ok: true, app: 'proctor-sheet', version: 1 });
}

function doPost(e) {
  var ss = null;
  try {
    var req = JSON.parse(e.postData.contents);
    if (SHARED_TOKEN && req.token !== SHARED_TOKEN) {
      return json_({ ok: false, error: '연결 토큰이 맞지 않습니다. 앱 설정의 토큰과 스크립트의 SHARED_TOKEN을 같게 하세요.' });
    }
    var book = req.book;
    ss = SpreadsheetApp.create(book.title);
    var first = ss.getSheets()[0];
    book.sheets.forEach(function (spec, i) {
      var sh = i === 0 ? first : ss.insertSheet();
      sh.setName(spec.name);
      render_(sh, spec);
    });
    var file = DriveApp.getFileById(ss.getId());
    if (req.folderId) file.moveTo(DriveApp.getFolderById(req.folderId));
    if (req.shareLink) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return json_({ ok: true, url: ss.getUrl(), id: ss.getId() });
  } catch (err) {
    // 도중에 실패하면 반쯤 만든 시트를 휴지통으로 옮긴다(복구 가능)
    if (ss) {
      try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch (ignore) {}
    }
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function render_(sh, spec) {
  var R = Math.max(spec.rows.length, 1);
  var C = 1;
  spec.rows.forEach(function (row) { C = Math.max(C, row.length); });
  if (sh.getMaxColumns() < C) sh.insertColumnsAfter(sh.getMaxColumns(), C - sh.getMaxColumns());
  if (sh.getMaxRows() < R) sh.insertRowsAfter(sh.getMaxRows(), R - sh.getMaxRows());

  var values = [], bgs = [], weights = [], colors = [], aligns = [], wraps = [], sizes = [];
  for (var r = 0; r < R; r++) {
    var vr = [], br = [], wr = [], cr = [], ar = [], xr = [], sr = [];
    for (var c = 0; c < C; c++) {
      var cell = (spec.rows[r] || [])[c] || null;
      vr.push(cell && cell.v !== null && cell.v !== undefined ? cell.v : '');
      br.push(cell && cell.bg ? cell.bg : null);
      wr.push(cell && cell.b ? 'bold' : 'normal');
      cr.push(cell && cell.fc ? cell.fc : '#000000');
      ar.push(cell && cell.al ? cell.al : 'center');
      xr.push(cell ? cell.wrap !== false : false);
      sr.push(cell && cell.fs ? cell.fs : 10);
    }
    values.push(vr); bgs.push(br); weights.push(wr); colors.push(cr); aligns.push(ar); wraps.push(xr); sizes.push(sr);
  }
  var range = sh.getRange(1, 1, R, C);
  range.setNumberFormat('@');
  range.setValues(values);
  range.setBackgrounds(bgs);
  range.setFontWeights(weights);
  range.setFontColors(colors);
  range.setHorizontalAlignments(aligns);
  range.setWraps(wraps);
  range.setFontSizes(sizes);
  range.setVerticalAlignment('middle');

  (spec.borders || []).forEach(function (b) {
    sh.getRange(b[0], b[1], b[2] - b[0] + 1, b[3] - b[1] + 1)
      .setBorder(true, true, true, true, true, true, '#c9d1e3', SpreadsheetApp.BorderStyle.SOLID);
  });
  (spec.merges || []).forEach(function (m) {
    sh.getRange(m[0], m[1], m[2] - m[0] + 1, m[3] - m[1] + 1).merge();
  });
  (spec.widths || []).forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  var hs = spec.heights || {};
  Object.keys(hs).forEach(function (k) { sh.setRowHeight(Number(k), hs[k]); });
  // 고정은 보기 편의 기능이므로 실패해도 시트 만들기는 계속한다
  try { if (spec.frozenRows) sh.setFrozenRows(spec.frozenRows); } catch (ignore) {}
  try { if (spec.frozenCols) sh.setFrozenColumns(spec.frozenCols); } catch (ignore) {}
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
`;
}
