import ExcelJS from 'exceljs';
import type { Cell, Day, Project, Role, Room, Slot, Teacher } from './types';
import { cellKey, computeStats, docTitle, emptyProject, fmtDay, fmtMD, isAssigned, needKey, roleColor, uid, type Stats } from './model';
import { normalize } from './store';
import { download, safeFileName } from './ui';

const BRAND = '4051E9';
const LINE = 'DDE2EE';
const DAYLINE = '8F9BEF';
const argb = (hex: string) => 'FF' + hex.replace('#', '').toUpperCase();
const solid = (hex: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: argb(hex) } });
const center: Partial<ExcelJS.Alignment> = { horizontal: 'center', vertical: 'middle', wrapText: true };
const thin = (hex = LINE): Partial<ExcelJS.Border> => ({ style: 'thin', color: { argb: argb(hex) } });

function box(c: ExcelJS.Cell) {
  c.border = { top: thin(), left: thin(), bottom: thin(), right: thin() };
  c.alignment = center;
}
function head(c: ExcelJS.Cell, v: string) {
  c.value = v;
  c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  c.fill = solid(BRAND);
  box(c);
}

type Display = 'name' | 'number';

function buildBuckets(p: Project, slots: Slot[], display: Display) {
  const map = new Map<string, string[]>();
  p.teachers.forEach((t, ti) =>
    slots.forEach((s) => {
      const c = p.cells[cellKey(t.id, s.key)];
      if (!isAssigned(c) || !c!.role || !c!.room) return;
      const k = `${s.key}|${c!.role}|${c!.room}`;
      if (!map.has(k)) map.set(k, []);
      map.get(k)!.push(display === 'name' ? t.name : String(ti + 1));
    }),
  );
  return map;
}

function visibleRoles(p: Project, s: Slot, hideEmpty: boolean): { r: Role; ri: number }[] {
  const all = p.roles.map((r, ri) => ({ r, ri }));
  if (!hideEmpty) return all;
  const vis = all.filter(({ r }) => p.rooms.some((m) => (p.need[needKey(s.key, r.id, m.id)] || 0) > 0));
  return vis.length ? vis : all;
}

function writeChart(ws: ExcelJS.Worksheet, p: Project, slots: Slot[], bk: Map<string, string[]>, startRow: number, title: string): number {
  const M = p.rooms.length;
  const cols = 3 + M + 1;
  ws.mergeCells(startRow, 1, startRow, cols);
  const tc = ws.getCell(startRow, 1);
  tc.value = title;
  tc.font = { bold: true, size: 15 };
  tc.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(startRow).height = 30;
  const hr = startRow + 1;
  head(ws.getCell(hr, 1), '날짜');
  head(ws.getCell(hr, 2), '교시');
  head(ws.getCell(hr, 3), '보직');
  p.rooms.forEach((m, i) => head(ws.getCell(hr, 4 + i), m.name));
  head(ws.getCell(hr, cols), '합계');
  ws.getRow(hr).height = 22;

  let r = hr + 1;
  let i = 0;
  while (i < slots.length) {
    const dayIdx = slots[i].dayIdx;
    const dayStart = r;
    for (; i < slots.length && slots[i].dayIdx === dayIdx; i++) {
      const s = slots[i];
      const slotStart = r;
      for (const { r: role, ri } of visibleRoles(p, s, true)) {
        let lines = 1;
        let sum = 0;
        ws.getCell(r, 1).value = fmtDay(s.date);
        ws.getCell(r, 2).value = `${s.period}교시`;
        const rc = ws.getCell(r, 3);
        rc.value = role.name;
        rc.fill = solid(roleColor(ri));
        rc.font = { bold: true };
        p.rooms.forEach((m, mi) => {
          const names = bk.get(`${s.key}|${role.id}|${m.id}`) || [];
          sum += p.need[needKey(s.key, role.id, m.id)] || 0;
          lines = Math.max(lines, names.length);
          const c = ws.getCell(r, 4 + mi);
          c.value = names.join('\n');
          box(c);
        });
        ws.getCell(r, cols).value = sum || '';
        for (let c = 1; c <= cols; c++) box(ws.getCell(r, c));
        ws.getRow(r).height = Math.max(20, lines * 15 + 4);
        r++;
      }
      if (r - 1 > slotStart) ws.mergeCells(slotStart, 2, r - 1, 2);
    }
    if (r - 1 > dayStart) ws.mergeCells(dayStart, 1, r - 1, 1);
    ws.getCell(dayStart, 1).font = { bold: true };
    for (let c = 1; c <= cols; c++) {
      const cell = ws.getCell(r - 1, c);
      cell.border = { ...cell.border, bottom: { style: 'medium', color: { argb: argb(DAYLINE) } } };
    }
  }
  ws.getColumn(1).width = 11;
  ws.getColumn(2).width = 7;
  ws.getColumn(3).width = 10;
  for (let c = 4; c < cols; c++) ws.getColumn(c).width = 9;
  ws.getColumn(cols).width = 6;
  return r;
}

const pageSetup = (orientation: 'landscape' | 'portrait'): Partial<ExcelJS.PageSetup> => ({
  paperSize: 9,
  orientation,
  fitToPage: true,
  fitToWidth: 1,
  fitToHeight: 0,
  margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
});

export async function exportWorkbook(p: Project, slots: Slot[], stats: Stats, display: Display = 'name') {
  const wb = new ExcelJS.Workbook();
  wb.creator = '정기고사 시감표';
  wb.created = new Date();
  const title = docTitle(p);
  const bk = buildBuckets(p, slots, display);
  const days = [...new Map(slots.map((s) => [s.dayIdx, s])).values()];

  // 1. 감독표
  {
    const ws = wb.addWorksheet('감독표', { pageSetup: pageSetup('landscape'), views: [{ showGridLines: false }] });
    writeChart(ws, p, slots, bk, 1, `${title} 감독표`.trim());
    ws.pageSetup.printTitlesRow = '2:2';
  }
  // 2. 감독표(일별)
  {
    const ws = wb.addWorksheet('감독표(일별)', { pageSetup: pageSetup('landscape'), views: [{ showGridLines: false }] });
    let row = 1;
    days.forEach((d, i) => {
      const next = writeChart(ws, p, slots.filter((s) => s.dayIdx === d.dayIdx), bk, row, `${title} 감독표 — ${fmtDay(d.date)}`.trim());
      if (i < days.length - 1) ws.getRow(next - 1).addPageBreak();
      row = next + 1;
    });
  }
  // 3. 개인시간표
  {
    const ws = wb.addWorksheet('개인시간표', { pageSetup: pageSetup('portrait'), views: [{ showGridLines: false }] });
    const roleById = new Map(p.roles.map((r, i) => [r.id, { r, i }]));
    const roomById = new Map(p.rooms.map((m) => [m.id, m.name]));
    let row = 1;
    p.teachers.forEach((t, ti) => {
      ws.mergeCells(row, 1, row, 4);
      const tc = ws.getCell(row, 1);
      tc.value = `${t.name} 선생님 감독 시간표`;
      tc.font = { bold: true, size: 15 };
      tc.alignment = { horizontal: 'center' };
      ws.getRow(row).height = 28;
      ws.mergeCells(row + 1, 1, row + 1, 4);
      const st = stats.teachers[ti];
      ws.getCell(row + 1, 1).value = `${title}  ·  총 ${st?.total ?? 0}시간`;
      ws.getCell(row + 1, 1).alignment = { horizontal: 'center' };
      ['날짜', '교시', '고사실', '보직'].forEach((h, i) => head(ws.getCell(row + 2, i + 1), h));
      row += 3;
      slots.forEach((s, si) => {
        const c = p.cells[cellKey(t.id, s.key)];
        const assigned = isAssigned(c);
        const role = c?.role ? roleById.get(c.role) : undefined;
        const first = si === 0 || slots[si - 1].dayIdx !== s.dayIdx;
        ws.getCell(row, 1).value = first ? fmtDay(s.date) : '';
        ws.getCell(row, 2).value = `${s.period}교시`;
        ws.getCell(row, 3).value = assigned && c?.room ? (roomById.get(c.room) ?? '') : '';
        const rc = ws.getCell(row, 4);
        rc.value = assigned ? (role?.r.name ?? '배정') : '';
        for (let k = 1; k <= 4; k++) box(ws.getCell(row, k));
        if (assigned && role) rc.fill = solid(roleColor(role.i));
        if (assigned) ws.getCell(row, 3).font = { bold: true };
        if (first && si > 0) for (let k = 1; k <= 4; k++) ws.getCell(row, k).border = { ...ws.getCell(row, k).border, top: { style: 'medium', color: { argb: argb(DAYLINE) } } };
        row++;
      });
      if (ti < p.teachers.length - 1) ws.getRow(row - 1).addPageBreak();
      row += 1;
    });
    [14, 10, 16, 16].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  }
  // 4. 감독현황
  {
    const ws = wb.addWorksheet('감독현황', { pageSetup: pageSetup('landscape'), views: [{ state: 'frozen', ySplit: 1, xSplit: 2 }] });
    const headers = ['순번', '이름', '배정할 시간', '총 감독 시간', ...days.map((d) => fmtDay(d.date)), '최대 연속', ...p.roles.map((r) => r.name), '현재 업무강도', '이전 누적', '누적 업무강도', '못 들어가는 고사실'];
    headers.forEach((h, i) => head(ws.getCell(1, i + 1), h));
    p.teachers.forEach((t, ti) => {
      const st = stats.teachers[ti];
      const vals = [ti + 1, t.name, t.target ?? '', st.total, ...days.map((d) => st.dayCounts[d.dayIdx] || 0), st.maxConsec, ...st.roleCounts, st.curLoad, t.prevLoad || 0, st.cumLoad, t.forbidden.map((id) => p.rooms.find((m) => m.id === id)?.name).filter(Boolean).join(', ')];
      vals.forEach((v, i) => {
        const c = ws.getCell(ti + 2, i + 1);
        c.value = v as ExcelJS.CellValue;
        box(c);
      });
    });
    headers.forEach((h, i) => (ws.getColumn(i + 1).width = Math.max(8, h.length * 2 + 2)));
  }
  // 5~7. proctor.xlsm 호환 시트
  writeCompatSheets(wb, p, slots, stats);

  // 원본 데이터(앱에서 다시 불러올 때 완전 복원용)
  {
    const ws = wb.addWorksheet('_data', { state: 'veryHidden' });
    const json = JSON.stringify(p);
    for (let i = 0, r = 1; i < json.length; i += 30000, r++) ws.getCell(r, 1).value = json.slice(i, i + 30000);
  }

  const buf = await wb.xlsx.writeBuffer();
  download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${safeFileName(title || '시감표')}_감독표.xlsx`);
}

function writeCompatSheets(wb: ExcelJS.Workbook, p: Project, slots: Slot[], stats: Stats) {
  const days = [...new Map(slots.map((s) => [s.dayIdx, s])).values()];
  const dayOrder = days.map((d) => p.days.find((x) => x.id === d.dayId)!);

  // 감독배정
  {
    const ws = wb.addWorksheet('감독배정', { views: [{ state: 'frozen', xSplit: 5, ySplit: 6 }] });
    const S = slots.length;
    ['순번', '이름', '배정할\n시간', '못들어가는\n고사실'].forEach((h, i) => head(ws.getCell(6, 2 + i), h));
    slots.forEach((s, si) => {
      ws.getCell(4, 6 + si).value = s.dayIdx + 1;
      ws.getCell(4, 6 + si).font = { color: { argb: 'FFFFFFFF' } };
      ws.getCell(5, 6 + si).value = s.firstOfDay ? fmtMD(s.date) : '';
      head(ws.getCell(6, 6 + si), '');
      ws.getCell(6, 6 + si).value = s.period;
      ws.getColumn(6 + si).width = 8;
    });
    const tail = ['최대\n가능시간', '총감독\n시간', '누적\n업무강도', ...p.roles.map((r) => r.name), '현재고사\n업무강도'];
    tail.forEach((h, i) => head(ws.getCell(6, 6 + S + i), h));
    ws.getCell(5, 6 + S + 3).value = '보직별 감독수';
    ws.getRow(6).height = 32;
    const roleIdx = new Map(p.roles.map((r, i) => [r.id, i]));
    const roomName = new Map(p.rooms.map((m) => [m.id, m.name]));
    p.teachers.forEach((t, ti) => {
      const r = 7 + ti;
      const st = stats.teachers[ti];
      ws.getCell(r, 2).value = ti + 1;
      ws.getCell(r, 3).value = t.name;
      ws.getCell(r, 4).value = t.target ?? null;
      ws.getCell(r, 5).value = t.forbidden.map((id) => roomName.get(id)).filter(Boolean).join(',');
      for (let c = 2; c <= 5; c++) box(ws.getCell(r, c));
      slots.forEach((s, si) => {
        const c = p.cells[cellKey(t.id, s.key)];
        const cell = ws.getCell(r, 6 + si);
        box(cell);
        if (!c) return;
        if (c.x) {
          cell.value = 'x';
          cell.fill = solid('#4B9461');
          cell.font = { color: { argb: 'FFFFFFFF' } };
          return;
        }
        const ri = c.role ? (roleIdx.get(c.role) ?? -1) : -1;
        cell.value = ri >= 0 && c.room ? `${roomName.get(c.room)}[${ri + 1}]` : isAssigned(c) ? 1 : null;
        if (c.fixed) cell.fill = solid('#E7E7E7');
        else if (ri >= 0) cell.fill = solid(roleColor(ri));
      });
      const tailVals = [st.maxPossible, st.total, st.cumLoad, ...st.roleCounts, st.curLoad];
      tailVals.forEach((v, i) => {
        const c = ws.getCell(r, 6 + S + i);
        c.value = v;
        box(c);
      });
    });
    ws.getColumn(3).width = 10;
    ws.getColumn(5).width = 14;
  }

  // 배정감독수정보
  {
    const ws = wb.addWorksheet('배정감독수정보', { views: [{ state: 'frozen', xSplit: 4, ySplit: 9 }] });
    head(ws.getCell(9, 2), '날짜');
    head(ws.getCell(9, 3), '교시');
    head(ws.getCell(9, 4), '보직');
    p.rooms.forEach((m, i) => head(ws.getCell(9, 5 + i), m.name));
    head(ws.getCell(9, 5 + p.rooms.length), '합계');
    let r = 10;
    slots.forEach((s) =>
      p.roles.forEach((role) => {
        ws.getCell(r, 1).value = s.dayIdx + 1;
        ws.getCell(r, 1).font = { color: { argb: 'FFFFFFFF' } };
        ws.getCell(r, 2).value = fmtMD(s.date);
        ws.getCell(r, 3).value = s.period;
        ws.getCell(r, 4).value = role.name;
        let sum = 0;
        p.rooms.forEach((m, i) => {
          const v = p.need[needKey(s.key, role.id, m.id)] || 0;
          sum += v;
          ws.getCell(r, 5 + i).value = v || null;
        });
        ws.getCell(r, 5 + p.rooms.length).value = sum;
        for (let c = 2; c <= 5 + p.rooms.length; c++) box(ws.getCell(r, c));
        r++;
      }),
    );
  }

  // 시험기본정보
  {
    const ws = wb.addWorksheet('시험기본정보');
    ws.getCell(7, 2).value = '1. 고사 날짜 및 교시';
    ws.getCell(7, 7).value = '2. 고사실 정보';
    ws.getCell(7, 10).value = '3. 감독교사 정보';
    ws.getCell(7, 14).value = '4. 업무강도';
    [
      [2, '순번'], [3, '날짜'], [4, '시작교시'], [5, '종료교시'],
      [7, '순번'], [8, '고사실명'],
      [10, '순번'], [11, '이름'], [12, '이전시험까지\n누적 업무강도'],
      [14, '순번'], [15, '보직'], [16, '업무강도'],
    ].forEach(([c, h]) => head(ws.getCell(9, c as number), h as string));
    ws.getRow(9).height = 44;
    dayOrder.forEach((d, i) => {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d.date);
      ws.getCell(10 + i, 2).value = i + 1;
      const dc = ws.getCell(10 + i, 3);
      dc.value = m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : d.date;
      dc.numFmt = 'mm"월" dd"일"';
      ws.getCell(10 + i, 4).value = d.start;
      ws.getCell(10 + i, 5).value = d.end;
    });
    p.rooms.forEach((m, i) => {
      ws.getCell(10 + i, 7).value = i + 1;
      ws.getCell(10 + i, 8).value = m.name;
    });
    p.teachers.forEach((t, i) => {
      ws.getCell(10 + i, 10).value = i + 1;
      ws.getCell(10 + i, 11).value = t.name;
      ws.getCell(10 + i, 12).value = t.prevLoad || 0;
    });
    p.roles.forEach((r, i) => {
      ws.getCell(10 + i, 14).value = i + 1;
      ws.getCell(10 + i, 15).value = r.name;
      ws.getCell(10 + i, 16).value = r.weight;
    });
    [3, 12].forEach((c) => (ws.getColumn(c).width = 14));
  }
}

// ---------------------------------------------------------------- 불러오기

type Raw = ExcelJS.CellValue | undefined;

function cellVal(c: ExcelJS.Cell): Raw {
  const v = c.value as unknown;
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    const o = v as Record<string, unknown>;
    if ('result' in o) return o.result as Raw;
    if ('richText' in o) return (o.richText as { text: string }[]).map((t) => t.text).join('');
    if ('text' in o) return o.text as Raw;
    if ('error' in o) return null;
  }
  return v as Raw;
}
const isEmpty = (v: Raw) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
const str = (v: Raw) => (v === null || v === undefined ? '' : v instanceof Date ? '' : String(v).trim());
const num = (v: Raw) => {
  if (typeof v === 'number') return v;
  const s = str(v);
  if (!s) return NaN;
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
};
const pad = (n: number) => String(n).padStart(2, '0');

function toISODate(v: Raw, fallbackYear: number): string {
  if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  if (typeof v === 'number') {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }
  const s = str(v);
  let m = /(\d{4})\D+(\d{1,2})\D+(\d{1,2})/.exec(s);
  if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
  m = /(\d{1,2})\D+(\d{1,2})/.exec(s);
  if (m) return `${fallbackYear}-${pad(+m[1])}-${pad(+m[2])}`;
  return '';
}

export async function importWorkbook(file: File): Promise<{ project: Project; notes: string[] }> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  const notes: string[] = [];

  const data = wb.getWorksheet('_data');
  if (data) {
    let json = '';
    for (let r = 1; ; r++) {
      const v = cellVal(data.getCell(r, 1));
      if (isEmpty(v)) break;
      json += String(v);
    }
    try {
      const p = normalize(JSON.parse(json));
      notes.push('이 앱에서 내보낸 파일입니다. 모든 입력과 배정 결과를 그대로 복원합니다.');
      notes.push(`일정 ${p.days.length}일 · 고사실 ${p.rooms.length}개 · 교사 ${p.teachers.length}명 · 보직 ${p.roles.length}개`);
      return { project: p, notes };
    } catch {
      notes.push('내장 데이터를 읽지 못해 시트 내용으로 불러옵니다.');
    }
  }

  const info = wb.getWorksheet('시험기본정보');
  if (!info) throw new Error("'시험기본정보' 시트가 없습니다. proctor 엑셀 파일이나 이 앱에서 내보낸 파일을 선택하세요.");
  const year = new Date().getFullYear();

  const days: Day[] = [];
  for (let r = 10; r < 1000; r++) {
    const dv = cellVal(info.getCell(r, 3));
    if (isEmpty(dv)) break;
    const start = num(cellVal(info.getCell(r, 4)));
    const end = num(cellVal(info.getCell(r, 5)));
    const s = Number.isFinite(start) ? start : 1;
    days.push({ id: uid(), date: toISODate(dv, year), start: s, end: Number.isFinite(end) ? end : s });
  }
  const rooms: Room[] = [];
  for (let r = 10; r < 2000; r++) {
    const v = str(cellVal(info.getCell(r, 8)));
    if (!v) break;
    rooms.push({ id: uid(), name: v });
  }
  const teachers: Teacher[] = [];
  for (let r = 10; r < 2000; r++) {
    const v = str(cellVal(info.getCell(r, 11)));
    if (!v) break;
    const load = num(cellVal(info.getCell(r, 12)));
    teachers.push({ id: uid(), name: v, prevLoad: Number.isFinite(load) ? load : 0, target: null, forbidden: [] });
  }
  const roles: Role[] = [];
  for (let r = 10; r < 200; r++) {
    const v = str(cellVal(info.getCell(r, 15)));
    if (!v) break;
    const w = num(cellVal(info.getCell(r, 16)));
    roles.push({ id: uid(), name: v, weight: Number.isFinite(w) ? w : 0 });
  }
  notes.push(`일정 ${days.length}일 · 고사실 ${rooms.length}개 · 교사 ${teachers.length}명 · 보직 ${roles.length}개`);

  const need: Record<string, number> = {};
  const needWs = wb.getWorksheet('배정감독수정보');
  if (needWs) {
    const roomCols = new Map<number, string>();
    for (let c = 5; c < 3000; c++) {
      const n = str(cellVal(needWs.getCell(9, c)));
      if (!n || n === '합계') break;
      const room = rooms.find((x) => x.name === n);
      if (room) roomCols.set(c, room.id);
    }
    let total = 0;
    for (let r = 10; r < 50000; r++) {
      const roleName = str(cellVal(needWs.getCell(r, 4)));
      if (!roleName) break;
      const day = days[num(cellVal(needWs.getCell(r, 1))) - 1];
      const per = num(cellVal(needWs.getCell(r, 3)));
      const role = roles.find((x) => x.name === roleName);
      if (!day || !role || !Number.isFinite(per)) continue;
      for (const [c, roomId] of roomCols) {
        const v = num(cellVal(needWs.getCell(r, c)));
        if (v > 0) {
          need[needKey(`${day.id}:${per}`, role.id, roomId)] = Math.floor(v);
          total += Math.floor(v);
        }
      }
    }
    notes.push(total ? `필요 감독 수 합계 ${total}시간` : '필요 감독 수가 비어 있습니다.');
  }

  const cells: Record<string, Cell> = {};
  const asg = wb.getWorksheet('감독배정');
  if (asg) {
    const colSlot = new Map<number, string>();
    for (let c = 6; c < 1000; c++) {
      const di = num(cellVal(asg.getCell(4, c)));
      const per = num(cellVal(asg.getCell(6, c)));
      if (!Number.isFinite(di) || !Number.isFinite(per)) break;
      const day = days[di - 1];
      if (day) colSlot.set(c, `${day.id}:${per}`);
    }
    let xCount = 0;
    let fixedCount = 0;
    let resultCount = 0;
    for (let r = 7; r < 5000; r++) {
      const name = str(cellVal(asg.getCell(r, 3)));
      if (!name) break;
      const t = teachers.find((x) => x.name === name);
      if (!t) continue;
      const tv = cellVal(asg.getCell(r, 4));
      if (!isEmpty(tv) && Number.isFinite(num(tv))) t.target = num(tv);
      const fb = str(cellVal(asg.getCell(r, 5)));
      if (fb)
        t.forbidden = fb
          .split(',')
          .map((s) => rooms.find((x) => x.name === s.trim())?.id)
          .filter((x): x is string => !!x);
      for (const [c, slotKey] of colSlot) {
        const cell = asg.getCell(r, c);
        const v = cellVal(cell);
        const s = str(v);
        const fill = cell.fill as ExcelJS.FillPattern | undefined;
        const gray = (fill?.fgColor?.argb || '').toUpperCase().endsWith('E7E7E7');
        let cc: Cell | null = null;
        const m = /^(.*)\[(\d+)\]$/.exec(s);
        if (s.toLowerCase() === 'x') {
          cc = { x: true };
          xCount++;
        } else if (m) {
          cc = gray ? { fixed: true } : { on: true };
          const room = rooms.find((x) => x.name === m[1]);
          const role = roles[Number(m[2]) - 1];
          if (room && role) {
            cc.role = role.id;
            cc.room = room.id;
            resultCount++;
          }
        } else if (num(v) === 1) cc = gray ? { fixed: true } : { on: true };
        if (cc?.fixed) fixedCount++;
        if (cc) cells[cellKey(t.id, slotKey)] = cc;
      }
    }
    notes.push(`감독 불가 ${xCount}칸 · 고정 ${fixedCount}칸${resultCount ? ` · 배정 결과 ${resultCount}칸` : ''}`);
  }

  const project = normalize({ ...emptyProject(), days, rooms, roles, teachers, need, cells });
  const st = computeStats(project, []);
  void st;
  notes.push('학교명과 고사 이름은 [기본 정보]에서 입력하세요.');
  return { project, notes };
}
