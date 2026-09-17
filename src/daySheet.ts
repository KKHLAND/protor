/**
 * 시험 당일 공유용 시감표(구글 시트 / 엑셀 공통 구조).
 * 학교 기존 양식("2026학년도 1학기 기말고사 시감표_1일차")을 따라 4개 시트를 만든다.
 *   1. 시감표      — 교시별 학년 시험 과목, 학급(고사실)별 정감독·부감독, 대표 교사
 *   2. 교사별      — 교사마다 교시별 정감독·부감독 칸에 고사실
 *   3. 감독 누계   — 그날까지 누계, 그날 소계, 보직별 누계
 *   4. 시험 시간표 — 불러온 시험 시간표가 있을 때
 * 세 번째 이후 보직(대기실 감독 등)은 정감독 칸에, 이미 차 있으면 부감독 칸에 적는다.
 */
import type { Project, Slot } from './types';
import { cellKey, isAssigned, parseDate } from './model';

export interface SheetCell {
  v: string | number | null;
  b?: boolean; // 굵게
  bg?: string; // 배경색
  fc?: string; // 글자색
  al?: 'left' | 'center' | 'right';
  wrap?: boolean;
  fs?: number; // 글자 크기(pt)
}

export interface SheetSpec {
  name: string;
  rows: (SheetCell | null)[][];
  merges: [number, number, number, number][]; // [r1, c1, r2, c2] (1부터)
  widths: number[]; // 픽셀
  heights?: Record<number, number>; // 행 번호 → 픽셀
  borders?: [number, number, number, number][];
  frozenRows?: number;
  frozenCols?: number;
}

export interface SheetBook {
  title: string;
  sheets: SheetSpec[];
}

const COLOR = { head: '#E3E6FF', label: '#F4F6FB', wait: '#FFF4E0', today: '#FFE8A3', title: '#1D1D1D' };
const WEEK = ['일', '월', '화', '수', '목', '금', '토'];

class Grid {
  rows: (SheetCell | null)[][] = [];
  merges: [number, number, number, number][] = [];
  borders: [number, number, number, number][] = [];
  heights: Record<number, number> = {};
  set(r: number, c: number, cell: SheetCell) {
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push(null);
    row[c - 1] = cell;
  }
  merge(r1: number, c1: number, r2: number, c2: number) {
    if (r1 !== r2 || c1 !== c2) this.merges.push([r1, c1, r2, c2]);
  }
  spec(name: string, widths: number[], extra: Partial<SheetSpec> = {}): SheetSpec {
    const s: SheetSpec = { name, rows: this.rows, merges: this.merges, widths, heights: this.heights, borders: this.borders, ...extra };
    // 구글 시트는 병합 셀을 가로지르는 행·열 고정을 허용하지 않으므로 그런 고정은 뺀다
    if (s.frozenCols && crossesFreeze(s.merges, 'col', s.frozenCols)) delete s.frozenCols;
    if (s.frozenRows && crossesFreeze(s.merges, 'row', s.frozenRows)) delete s.frozenRows;
    return s;
  }
}

/** 고정 경계(n번째 행/열 뒤)를 가로지르는 병합이 있는지 */
export function crossesFreeze(merges: SheetSpec['merges'], kind: 'row' | 'col', n: number) {
  return merges.some(([r1, c1, r2, c2]) => (kind === 'col' ? c1 <= n && c2 > n : r1 <= n && r2 > n));
}

const gradeNo = (g: string) => Number(/(\d)/.exec(g)?.[1] ?? 9);

export function koreanDate(date: string, withWeekday = true) {
  const d = parseDate(date);
  if (!d) return date;
  return `${d.getMonth() + 1}월 ${d.getDate()}일${withWeekday ? `(${WEEK[d.getDay()]})` : ''}`;
}

/** 시험 과목 글에서 대표 교사 표기를 만든다 — "화법과 언어 (40분/공통) (8:20~9:00) 교사01, 교사02" → "화법과 언어 교사01" */
export function representative(examText: string): string {
  const m = /^(.*?)\s*\(\s*\d+\s*분/.exec(examText);
  const subjectsRaw = (m ? m[1] : examText.split('(')[0]).trim();
  const subjects = subjectsRaw
    .split('/')
    .map((s) => s.replace(/\([^)]*\)/g, '').trim())
    .filter(Boolean);
  const times = [...examText.matchAll(/\([^()]*\d{1,2}:\d{2}[^()]*\)/g)];
  const last = times[times.length - 1];
  const teacherRaw = last ? examText.slice((last.index ?? 0) + last[0].length).trim() : '';
  const groups = teacherRaw
    .split('/')
    .map((s) => s.trim())
    .filter(Boolean);
  const firstOf = (g: string) => g.split(',')[0].replace(/\s+/g, ' ').trim();
  if (subjects.length > 1 && groups.length === subjects.length) return subjects.map((s, i) => `${s} ${firstOf(groups[i])}`).join(', ');
  const first = groups[0] ? firstOf(groups[0]) : '';
  return `${subjects.join(' / ')}${first ? ` ${first}` : ''}`;
}

function examsFor(p: Project, slotKey: string) {
  const out: { grade: string; text: string }[] = [];
  for (const [k, v] of Object.entries(p.exams ?? {})) {
    const i = k.lastIndexOf('|');
    if (k.slice(0, i) === slotKey && v) out.push({ grade: k.slice(i + 1), text: v });
  }
  return out.sort((a, b) => gradeNo(a.grade) - gradeNo(b.grade));
}

export function dayBookTitle(p: Project, dayIdx: number) {
  return `${p.meta.title.trim() || '정기고사'} 시감표_${dayIdx + 1}일차`;
}

export function buildDayBook(p: Project, slots: Slot[], dayIdx: number): SheetBook {
  const daySlots = slots.filter((s) => s.dayIdx === dayIdx);
  if (!daySlots.length) throw new Error('선택한 날짜의 시험 시간이 없습니다.');
  const dayNo = dayIdx + 1;
  const exam = p.meta.title.trim() || '정기고사';
  const heading = `${exam} ${dayNo}일차(${koreanDate(daySlots[0].date, false)})`;
  const P = daySlots.length;
  const T = p.teachers.length;
  const roleIdx = new Map(p.roles.map((r, i) => [r.id, i]));
  const roomName = new Map(p.rooms.map((r) => [r.id, r.name]));
  const roleLabel = (i: number) => p.roles[i]?.name ?? (i === 0 ? '정감독' : '부감독');

  // 그날 배정을 (교시, 고사실) → [정감독 칸, 부감독 칸] 으로 모은다
  const items: { teacherId: string; name: string; slot: Slot; ri: number; room: string }[] = [];
  p.teachers.forEach((t) =>
    daySlots.forEach((s) => {
      const c = p.cells[cellKey(t.id, s.key)];
      if (!isAssigned(c) || !c?.role || !c.room) return;
      const ri = roleIdx.get(c.role);
      if (ri === undefined) return;
      items.push({ teacherId: t.id, name: t.name, slot: s, ri, room: c.room });
    }),
  );
  items.sort((a, b) => Number(a.ri > 1) - Number(b.ri > 1));
  const byRoom = new Map<string, [string[], string[]]>();
  const byTeacher = new Map<string, Map<string, { col: 0 | 1; room: string }>>();
  for (const it of items) {
    const key = `${it.slot.key}|${it.room}`;
    const entry = byRoom.get(key) ?? [[], []];
    let col: 0 | 1 = it.ri === 1 ? 1 : 0;
    if (it.ri > 1) col = entry[0].length === 0 ? 0 : entry[1].length === 0 ? 1 : 0;
    entry[col].push(it.name);
    byRoom.set(key, entry);
    if (!byTeacher.has(it.teacherId)) byTeacher.set(it.teacherId, new Map());
    byTeacher.get(it.teacherId)!.set(it.slot.key, { col, room: roomName.get(it.room) ?? '' });
  }

  // ---------- 1. 시감표 ----------
  const g = new Grid();
  const W = 1 + 2 * P;
  g.set(1, 1, { v: heading, b: true, fs: 15, fc: COLOR.title });
  g.merge(1, 1, 1, W);
  g.heights[1] = 40;
  let r = 3;
  const tableTop = r;
  g.set(r, 1, { v: '구분', b: true, bg: COLOR.head });
  daySlots.forEach((s, i) => {
    g.set(r, 2 + 2 * i, { v: `${s.period}교시`, b: true, bg: COLOR.head });
    g.merge(r, 2 + 2 * i, r, 3 + 2 * i);
  });
  r++;
  const grades = [...new Set(daySlots.flatMap((s) => examsFor(p, s.key).map((e) => e.grade)))].sort((a, b) => gradeNo(a) - gradeNo(b));
  for (const grade of grades) {
    g.set(r, 1, { v: grade || '시험', b: true, bg: COLOR.label });
    daySlots.forEach((s, i) => {
      g.set(r, 2 + 2 * i, { v: p.exams?.[`${s.key}|${grade}`] ?? '', wrap: true, fs: 9 });
      g.merge(r, 2 + 2 * i, r, 3 + 2 * i);
    });
    g.heights[r] = 62;
    r++;
  }
  const roleHeader = r;
  g.set(r, 1, { v: '학급', b: true, bg: COLOR.head });
  daySlots.forEach((_, i) => {
    g.set(r, 2 + 2 * i, { v: roleLabel(0), b: true, bg: COLOR.head });
    g.set(r, 3 + 2 * i, { v: roleLabel(1), b: true, bg: COLOR.head });
  });
  r++;
  for (const room of p.rooms) {
    g.set(r, 1, { v: room.name, b: true, bg: /대기/.test(room.name) ? COLOR.wait : COLOR.label });
    daySlots.forEach((s, i) => {
      const e = byRoom.get(`${s.key}|${room.id}`);
      g.set(r, 2 + 2 * i, { v: e?.[0].join('\n') ?? '' });
      g.set(r, 3 + 2 * i, { v: e?.[1].join('\n') ?? '' });
    });
    const lines = Math.max(1, ...daySlots.map((s) => Math.max(...(byRoom.get(`${s.key}|${room.id}`)?.map((x) => x.length) ?? [1]))));
    if (lines > 1) g.heights[r] = 21 * lines;
    r++;
  }
  if (grades.length) {
    g.set(r, 1, { v: '대표 교사', b: true, bg: COLOR.label });
    let maxLines = 1;
    daySlots.forEach((s, i) => {
      const lines = examsFor(p, s.key).map((e) => `${e.grade} ${representative(e.text)}`.trim());
      maxLines = Math.max(maxLines, lines.length);
      g.set(r, 2 + 2 * i, { v: lines.join('\n'), wrap: true, fs: 9 });
      g.merge(r, 2 + 2 * i, r, 3 + 2 * i);
    });
    g.heights[r] = Math.max(34, 18 * maxLines + 8);
    r++;
  }
  g.borders.push([tableTop, 1, r - 1, W]);
  // 휴대폰으로 볼 때 위쪽 과목 줄이 화면을 차지하지 않도록 고정하지 않는다
  void roleHeader;
  const sheet1 = g.spec('시감표', [150, ...Array(2 * P).fill(84)]);

  // ---------- 2. 교사별 ----------
  const g2 = new Grid();
  g2.set(1, 1, { v: '순번', b: true, bg: COLOR.head });
  g2.merge(1, 1, 2, 1);
  g2.set(1, 2, { v: '이름', b: true, bg: COLOR.head });
  g2.merge(1, 2, 2, 2);
  g2.set(1, 3, { v: heading, b: true, bg: COLOR.head });
  g2.merge(1, 3, 1, 2 + 2 * P);
  daySlots.forEach((s, i) => {
    g2.set(2, 3 + 2 * i, { v: `${s.period}교시`, b: true, bg: COLOR.head });
    g2.merge(2, 3 + 2 * i, 2, 4 + 2 * i);
  });
  g2.set(3, 1, { v: '합계', b: true, bg: COLOR.label });
  g2.set(3, 2, { v: `${byTeacher.size} / ${T}`, b: true, bg: COLOR.label });
  daySlots.forEach((_, i) => {
    g2.set(3, 3 + 2 * i, { v: roleLabel(0), b: true, bg: COLOR.label });
    g2.set(3, 4 + 2 * i, { v: roleLabel(1), b: true, bg: COLOR.label });
  });
  p.teachers.forEach((t, ti) => {
    const rr = 4 + ti;
    const mine = byTeacher.get(t.id);
    g2.set(rr, 1, { v: ti + 1 });
    g2.set(rr, 2, { v: t.name, b: true });
    daySlots.forEach((s, i) => {
      const a = mine?.get(s.key);
      g2.set(rr, 3 + 2 * i, { v: a && a.col === 0 ? a.room : '' });
      g2.set(rr, 4 + 2 * i, { v: a && a.col === 1 ? a.room : '' });
    });
  });
  g2.borders.push([1, 1, 3 + T, 2 + 2 * P]);
  const sheet2 = g2.spec('교사별', [50, 110, ...Array(2 * P).fill(84)], { frozenRows: 3, frozenCols: 2 });

  // ---------- 3. 감독 누계 ----------
  const g3 = new Grid();
  const W3 = 4 + p.roles.length;
  g3.set(1, 1, { v: heading, b: true, fs: 13 });
  g3.merge(1, 1, 1, W3);
  g3.heights[1] = 32;
  ['순번', '이름', '전체 감독\n(누계)', `${dayNo}일차 감독\n(소계)`, ...p.roles.map((x) => `${x.name}\n(누계)`)].forEach((h, i) =>
    g3.set(2, i + 1, { v: h, b: true, bg: COLOR.head, wrap: true }),
  );
  g3.heights[2] = 42;
  const upTo = slots.filter((s) => s.dayIdx <= dayIdx);
  p.teachers.forEach((t, ti) => {
    let total = 0;
    let today = 0;
    const perRole = p.roles.map(() => 0);
    upTo.forEach((s) => {
      const c = p.cells[cellKey(t.id, s.key)];
      if (!isAssigned(c)) return;
      total++;
      if (s.dayIdx === dayIdx) today++;
      const ri = c?.role ? roleIdx.get(c.role) : undefined;
      if (ri !== undefined) perRole[ri]++;
    });
    [ti + 1, t.name, total || '', today || '', ...perRole.map((n) => n || '')].forEach((v, i) => g3.set(3 + ti, i + 1, { v, b: i === 1 }));
  });
  g3.borders.push([2, 1, 2 + T, W3]);
  const sheet3 = g3.spec('감독 누계', [50, 110, 90, 90, ...p.roles.map(() => 90)], { frozenRows: 2 });

  const sheets = [sheet1, sheet2, sheet3];

  // ---------- 4. 시험 시간표 ----------
  const allGrades = [...new Set(Object.keys(p.exams ?? {}).map((k) => k.slice(k.lastIndexOf('|') + 1)))].sort((a, b) => gradeNo(a) - gradeNo(b));
  if (allGrades.length) {
    const days = [...new Map(slots.map((s) => [s.dayIdx, s])).values()];
    const G = allGrades.length;
    const g4 = new Grid();
    const W4 = 1 + days.length * G;
    g4.set(1, 1, { v: `${exam} 시험 시간표`, b: true, fs: 13 });
    g4.merge(1, 1, 1, W4);
    g4.heights[1] = 32;
    g4.set(2, 1, { v: '교시', b: true, bg: COLOR.head });
    g4.merge(2, 1, 3, 1);
    days.forEach((d, i) => {
      const c = 2 + i * G;
      const bg = d.dayIdx === dayIdx ? COLOR.today : COLOR.head;
      g4.set(2, c, { v: `${d.dayIdx + 1}일차 ${koreanDate(d.date)}`, b: true, bg });
      g4.merge(2, c, 2, c + G - 1);
      allGrades.forEach((gr, j) => g4.set(3, c + j, { v: gr || '시험', b: true, bg }));
    });
    const maxPeriod = Math.max(...slots.map((s) => s.period));
    for (let per = 1; per <= maxPeriod; per++) {
      const rr = 3 + per;
      g4.set(rr, 1, { v: `${per}교시`, b: true, bg: COLOR.label });
      days.forEach((d, i) =>
        allGrades.forEach((gr, j) => g4.set(rr, 2 + i * G + j, { v: p.exams?.[`${d.dayId}:${per}|${gr}`] ?? '', wrap: true, fs: 9 })),
      );
      g4.heights[rr] = 76;
    }
    g4.borders.push([2, 1, 3 + maxPeriod, W4]);
    sheets.push(g4.spec('시험 시간표', [60, ...Array(days.length * G).fill(140)], { frozenRows: 3 }));
  }

  return { title: dayBookTitle(p, dayIdx), sheets };
}
