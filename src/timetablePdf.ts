/**
 * 학교 시험 시간표 PDF 읽기.
 * PDF에는 표 구조가 없고 글자 조각과 좌표만 있으므로
 *   1) 같은 줄의 가까운 조각을 이어 "덩어리(run)"로 만들고
 *   2) 날짜가 두 개 이상 있는 줄을 표 머리글로, 그 아래 "N학년"을 열, 왼쪽 "N교시"를 행으로 삼아
 *   3) 각 덩어리를 가장 가까운 열·교시 칸에 모은다.
 * 칸 안의 시험 시각(예: 8:20~9:00)과 본령 시각(1교시 08:20 …)이 있으면 교시를 시각으로 다시 확인한다.
 */
import type { ParsedTimetable, TimetableExam } from './timetable';

export interface PdfItem {
  str: string;
  x: number;
  y: number; // PDF 좌표: 위로 갈수록 큼
  w: number;
  h: number;
}

interface Run {
  text: string;
  x1: number;
  x2: number;
  y: number;
  h: number;
}

const cx = (r: Run) => (r.x1 + r.x2) / 2;
const pad = (n: number) => String(n).padStart(2, '0');
const DATE = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/;
const TIME_RANGE = /(\d{1,2}):(\d{2})\s*[~∼～〜-]\s*(\d{1,2}):(\d{2})/;

/** 같은 줄에서 가까운 글자 조각을 이어 붙인다 */
export function toRuns(items: PdfItem[]): Run[] {
  const sorted = items.filter((i) => i.str.trim()).sort((a, b) => b.y - a.y || a.x - b.x);
  const lines: PdfItem[][] = [];
  for (const it of sorted) {
    const line = lines.find((l) => Math.abs(l[0].y - it.y) <= Math.max(1.5, (it.h || 8) * 0.3));
    if (line) line.push(it);
    else lines.push([it]);
  }
  const runs: Run[] = [];
  for (const line of lines) {
    line.sort((a, b) => a.x - b.x);
    let cur: Run | null = null;
    for (const it of line) {
      const h = it.h || 8;
      const gap = cur ? it.x - cur.x2 : Infinity;
      if (cur && gap <= h * 0.9) {
        cur.text += gap > h * 0.15 ? ` ${it.str}` : it.str;
        cur.x2 = Math.max(cur.x2, it.x + it.w);
      } else {
        if (cur) runs.push(cur);
        cur = { text: it.str, x1: it.x, x2: it.x + it.w, y: it.y, h };
      }
    }
    if (cur) runs.push(cur);
  }
  return runs.map((r) => ({ ...r, text: r.text.replace(/\s+/g, ' ').trim() })).filter((r) => r.text);
}

const minutes = (h: string, m: string) => Number(h) * 60 + Number(m);

export function parseTimetableItems(pages: PdfItem[][]): ParsedTimetable | null {
  const allRuns = pages.map(toRuns);

  let title = '';
  let year = 0;
  for (const runs of allRuns) {
    for (const r of runs) {
      const m = /(\d{4})\s*학년도.*(중간|기말|고사|시험)/.exec(r.text);
      if (m) {
        year = Number(m[1]);
        title = r.text
          .replace(/\s*(시험\s*)?시간표\s*$/, '')
          .replace(/^\d+\s*[.)]\s*/, '')
          .replace(/\s+/g, ' ')
          .trim();
        break;
      }
    }
    if (title) break;
  }
  const secondSemester = /2\s*학기/.test(title);
  const isoOf = (text: string) => {
    const m = DATE.exec(text);
    if (!m) return null;
    let y = year || new Date().getFullYear();
    if (secondSemester && Number(m[1]) <= 2) y += 1;
    return `${y}-${pad(Number(m[1]))}-${pad(Number(m[2]))}`;
  };

  // key: date|period|grade → 덩어리 목록
  const buckets = new Map<string, { date: string; period: number; grade: string; runs: Run[]; timePeriod?: number }>();

  for (const runs of allRuns) {
    // 표 머리글: 날짜가 2개 이상 있는 줄
    const dateRuns = runs.filter((r) => isoOf(r.text) && r.text.length <= 20);
    const headerLines: Run[][] = [];
    for (const r of dateRuns.sort((a, b) => b.y - a.y)) {
      const line = headerLines.find((l) => Math.abs(l[0].y - r.y) <= 4);
      if (line) line.push(r);
      else headerLines.push([r]);
    }
    const headers = headerLines.filter((l) => l.length >= 2).sort((a, b) => b[0].y - a[0].y);

    headers.forEach((header, hi) => {
      const yTop = Math.max(...header.map((r) => r.y));
      const yBottom = hi + 1 < headers.length ? Math.max(...headers[hi + 1].map((r) => r.y)) : -Infinity;
      const dates = header.map((r) => ({ date: isoOf(r.text)!, x: cx(r) })).sort((a, b) => a.x - b.x);
      const block = runs.filter((r) => r.y > yBottom + 2 && r.y < yTop + 30);

      // 열: 머리글 근처의 "N학년" (없으면 날짜 자체). 첫 학년 열은 날짜 글자보다 왼쪽에 있을 수 있다
      const gradeRuns = block.filter(
        (r) => /^[1-6]\s*학년$/.test(r.text) && Math.abs(r.y - yTop) <= 30 && cx(r) >= dates[0].x - (dates[1] ? dates[1].x - dates[0].x : 120),
      );
      const columns =
        gradeRuns.length > 0
          ? gradeRuns.map((r) => {
              const x = cx(r);
              const date = dates.reduce((best, d) => (Math.abs(d.x - x) < Math.abs(best.x - x) ? d : best)).date;
              return { x, date, grade: `${/(\d)/.exec(r.text)![1]}학년` };
            })
          : dates.map((d) => ({ x: d.x, date: d.date, grade: '' }));
      columns.sort((a, b) => a.x - b.x);
      const gaps = columns.slice(1).map((c, i) => c.x - columns[i].x).sort((a, b) => a - b);
      const colWidth = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 120;
      // 표 데이터 영역의 왼쪽 경계 (이보다 왼쪽은 교시·본령 칸)
      const firstDateX = columns[0].x - colWidth / 2;

      // 교시 행
      const periodLabels: { period: number; y: number; start?: number }[] = [];
      for (const r of block) {
        if (cx(r) >= firstDateX) continue;
        const m = /^(\d{1,2})\s*교시(?:\s*0?(\d{1,2}):(\d{2}))?$/.exec(r.text);
        if (!m || r.y >= yTop + 2) continue;
        if (periodLabels.some((p) => p.period === Number(m[1]))) continue;
        periodLabels.push({ period: Number(m[1]), y: r.y, start: m[2] ? minutes(m[2], m[3]) : undefined });
      }
      if (!periodLabels.length) return;
      periodLabels.sort((a, b) => b.y - a.y);
      // 본령 시각 (교시 글자와 같은 줄 근처의 hh:mm)
      for (const r of block) {
        if (cx(r) >= firstDateX) continue;
        const m = /^0?(\d{1,2}):(\d{2})$/.exec(r.text);
        if (!m) continue;
        const near = periodLabels.reduce((best, p) => (Math.abs(p.y - r.y) < Math.abs(best.y - r.y) ? p : best));
        if (Math.abs(near.y - r.y) <= 14 && near.start === undefined) near.start = minutes(m[1], m[2]);
      }

      const headerY = Math.min(yTop, ...gradeRuns.map((r) => r.y));

      // 교시 행의 세로 범위
      const bands = periodLabels.map((p, i) => {
        const prev = periodLabels[i - 1];
        const next = periodLabels[i + 1];
        const half = next ? (p.y - next.y) / 2 : prev ? (prev.y - p.y) / 2 : 40;
        const top = prev ? (prev.y + p.y) / 2 : Math.min(headerY - 2, p.y + half);
        const bottom = next ? (p.y + next.y) / 2 : Math.max(yBottom + 2, p.y - half);
        return { ...p, top, bottom };
      });

      const skip = new Set<Run>([...header, ...gradeRuns]);
      for (const r of block) {
        if (skip.has(r)) continue;
        const x = cx(r);
        if (x < firstDateX) continue;
        if (r.y >= headerY - 1) continue;
        const col = columns.reduce((best, c) => (Math.abs(c.x - x) < Math.abs(best.x - x) ? c : best));
        if (Math.abs(col.x - x) > colWidth * 0.8) continue;
        const band = bands.find((b) => r.y <= b.top && r.y > b.bottom);
        if (!band) continue;
        const key = `${col.date}|${band.period}|${col.grade}`;
        if (!buckets.has(key)) buckets.set(key, { date: col.date, period: band.period, grade: col.grade, runs: [] });
        const bucket = buckets.get(key)!;
        bucket.runs.push(r);
        // 시험 시각으로 교시 다시 확인
        const tm = TIME_RANGE.exec(r.text);
        const withStart = bands.filter((b) => b.start !== undefined);
        if (tm && withStart.length && bucket.timePeriod === undefined) {
          const t = minutes(tm[1], tm[2]);
          const fit = withStart.filter((b) => b.start! <= t + 5).sort((a, b) => b.start! - a.start!)[0];
          if (fit) bucket.timePeriod = fit.period;
        }
      }
    });
  }

  const exams = new Map<string, TimetableExam>();
  for (const b of buckets.values()) {
    const text = b.runs
      .sort((a, c) => c.y - a.y || a.x1 - c.x1)
      .map((r) => r.text)
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (!text || /^[xX×-]$/.test(text)) continue;
    // 과목 칸에는 보통 시험 시각이 들어 있다. 시각이 없고 아주 짧은 글은 버린다
    if (!TIME_RANGE.test(text) && text.length < 4) continue;
    const period = b.timePeriod ?? b.period;
    const key = `${b.date}|${period}|${b.grade}`;
    const prev = exams.get(key);
    exams.set(key, { date: b.date, period, grade: b.grade, text: prev ? `${prev.text} ${text}` : text });
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
  return { title, sheetName: 'PDF', days, exams: [...exams.values()] };
}

/** 브라우저에서 PDF 파일을 읽어 시험 시간표로 변환 */
export async function parseTimetablePdf(file: File): Promise<ParsedTimetable | null> {
  const pdfjs = await import('pdfjs-dist');
  if (!pdfjs.GlobalWorkerOptions.workerPort) {
    const { default: PdfWorker } = await import('pdfjs-dist/build/pdf.worker.min.mjs?worker&inline');
    pdfjs.GlobalWorkerOptions.workerPort = new PdfWorker();
  }
  const doc = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pages: PdfItem[][] = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    pages.push(
      content.items
        .filter((it): it is Extract<typeof it, { str: string }> => 'str' in it)
        .map((it) => ({ str: it.str, x: it.transform[4], y: it.transform[5], w: it.width, h: it.height || Math.abs(it.transform[3]) })),
    );
  }
  return parseTimetableItems(pages);
}
