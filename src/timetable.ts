/**
 * 학교 시험 시간표(구글 시트를 xlsx로 받은 파일)에서 시험 날짜·교시·과목을 읽는다.
 *
 * 인식하는 모양
 *   교시 | 본령 | 10월 1일(목) | 10월 1일(목) | 10월 2일(금) …   ← "교시" 칸과 날짜 칸이 있는 머리글 줄
 *   교시 | 본령 | 2학년        | 3학년        | 2학년 …           ← (선택) 학년 줄
 *   1    | 8:20 | 화법과 언어 (40분/공통) (8:20~9:00) 교사01, 교사02 …
 * 머리글 블록이 여러 개(2·3학년 표, 1학년 표)여도 모두 합친다. 날짜 칸은 글자·날짜 서식 모두 읽는다.
 */
import type ExcelJS from 'exceljs';
import type { Day, Project } from './types';
import { cleanup, fmtDay, uid } from './model';

export interface TimetableExam {
  date: string;
  period: number;
  grade: string;
  text: string;
}

export interface ParsedTimetable {
  title: string;
  sheetName: string;
  days: { date: string; start: number; end: number; periods: number[] }[];
  exams: TimetableExam[];
}

const pad = (n: number) => String(n).padStart(2, '0');

function raw(cell: ExcelJS.Cell): unknown {
  const src = cell.isMerged && cell.master ? cell.master : cell;
  const v = src.value as unknown;
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    const o = v as Record<string, unknown>;
    if ('richText' in o) return (o.richText as { text: string }[]).map((t) => t.text).join('');
    if ('result' in o) return o.result;
    if ('text' in o) return o.text;
    return null;
  }
  return v;
}

const text = (v: unknown) => (v === null || v === undefined || v instanceof Date ? '' : String(v).replace(/\s+/g, ' ').trim());

function parseSheet(ws: ExcelJS.Worksheet): ParsedTimetable | null {
  const maxR = Math.min(ws.rowCount, 500);
  const maxC = Math.min(ws.columnCount, 100);
  const get = (r: number, c: number) => raw(ws.getCell(r, c));

  let title = '';
  let year = 0;
  for (let r = 1; r <= Math.min(maxR, 20) && !title; r++) {
    for (let c = 1; c <= maxC; c++) {
      const t = text(get(r, c));
      const m = /(\d{4})\s*학년도/.exec(t);
      if (m && /시간표|고사|시험/.test(t)) {
        year = Number(m[1]);
        title = t
          .replace(/시험\s*시간표|시간표/g, '')
          .replace(/\s+/g, ' ')
          .trim();
        break;
      }
    }
  }
  const secondSemester = /2\s*학기/.test(title);
  const dateOf = (v: unknown): string | null => {
    if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
    const m = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/.exec(text(v));
    if (!m) return null;
    let y = year || new Date().getFullYear();
    // 2학기 시험이 해를 넘겨 1~2월에 있으면 다음 해
    if (secondSemester && Number(m[1]) <= 2) y += 1;
    return `${y}-${pad(Number(m[1]))}-${pad(Number(m[2]))}`;
  };

  const exams = new Map<string, TimetableExam>();
  for (let r = 1; r <= maxR; r++) {
    let periodCol = 0;
    const colDate = new Map<number, string>();
    for (let c = 1; c <= maxC; c++) {
      const v = get(r, c);
      if (!periodCol && text(v).replace(/\s/g, '') === '교시') periodCol = c;
      const d = dateOf(v);
      if (d) colDate.set(c, d);
    }
    if (!periodCol || !colDate.size) continue;

    const colGrade = new Map<number, string>();
    for (const c of colDate.keys()) {
      const m = /([1-6])\s*학년/.exec(text(get(r + 1, c)));
      if (m) colGrade.set(c, `${m[1]}학년`);
    }
    const first = colGrade.size ? r + 2 : r + 1;
    let rr = first;
    for (; rr <= maxR; rr++) {
      const pm = /^(\d{1,2})\s*(교시)?$/.exec(text(get(rr, periodCol)));
      if (!pm) break;
      const period = Number(pm[1]);
      if (period < 1 || period > 12) break;
      for (const [c, date] of colDate) {
        const t = text(get(rr, c));
        if (!t || /^[xX×-]$/.test(t)) continue;
        const grade = colGrade.get(c) ?? '';
        const key = `${date}|${period}|${grade}`;
        if (!exams.has(key)) exams.set(key, { date, period, grade, text: t });
      }
    }
    r = Math.max(r, rr - 1);
  }
  if (!exams.size) return null;

  const byDate = new Map<string, Set<number>>();
  exams.forEach((e) => {
    if (!byDate.has(e.date)) byDate.set(e.date, new Set());
    byDate.get(e.date)!.add(e.period);
  });
  const days = [...byDate.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, set]) => {
      const periods = [...set].sort((a, b) => a - b);
      return { date, start: periods[0], end: periods[periods.length - 1], periods };
    });
  return { title, sheetName: ws.name, days, exams: [...exams.values()] };
}

/** 워크북의 시트 중 시험 시간표 모양인 첫 시트를 읽는다 */
export function parseTimetable(wb: ExcelJS.Workbook): ParsedTimetable | null {
  for (const ws of wb.worksheets) {
    const res = parseSheet(ws);
    if (res) return res;
  }
  return null;
}

/**
 * 읽은 시간표를 현재 작업에 적용한다. 같은 날짜가 이미 있으면 그 일정을 그대로 이어 써서
 * 입력해 둔 필요 감독 수·감독 불가 표시가 유지된다.
 */
export function applyTimetable(p: Project, t: ParsedTimetable): { project: Project; notes: string[] } {
  const byDate = new Map(p.days.map((d) => [d.date, d]));
  const days: Day[] = t.days.map((td) => ({ id: byDate.get(td.date)?.id ?? uid(), date: td.date, start: td.start, end: td.end }));
  const idByDate = new Map(days.map((d) => [d.date, d.id]));
  const exams: Record<string, string> = {};
  t.exams.forEach((e) => {
    const id = idByDate.get(e.date);
    if (id) exams[`${id}:${e.period}|${e.grade}`] = e.text;
  });

  const notes: string[] = [];
  if (t.title) notes.push(`고사 이름: ${t.title}`);
  notes.push(`시험 일정 ${days.length}일`);
  t.days.forEach((d, i) => {
    const gaps = [];
    for (let per = d.start; per <= d.end; per++) if (!d.periods.includes(per)) gaps.push(per);
    notes.push(`${i + 1}일차 ${fmtDay(d.date)} ${d.start}~${d.end}교시${gaps.length ? ` (시험 없는 ${gaps.join('·')}교시 포함)` : ''}`);
  });
  const grades = [...new Set(t.exams.map((e) => e.grade).filter(Boolean))].sort();
  notes.push(`시험 과목 ${t.exams.length}건${grades.length ? ` · ${grades.join(', ')}` : ''}`);
  const removed = p.days.filter((d) => !idByDate.has(d.date));
  if (removed.length) notes.push(`시간표에 없는 기존 날짜 ${removed.map((d) => fmtDay(d.date)).join(', ')}은(는) 삭제됩니다.`);
  const kept = p.days.length - removed.length;
  if (kept > 0) notes.push(`같은 날짜 ${kept}일은 이미 입력한 필요 감독 수·감독 불가 표시를 그대로 유지합니다.`);

  return {
    project: cleanup({ ...p, days, exams, meta: { ...p.meta, title: t.title || p.meta.title } }),
    notes,
  };
}
