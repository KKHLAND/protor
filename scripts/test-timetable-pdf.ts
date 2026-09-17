// 실제 학교 시험 시간표 PDF(../2026학년도 2학기 중간고사 시간표.pdf) 읽기 검증
import { existsSync, readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

const PDF = '../2026학년도 2학기 중간고사 시간표.pdf';
if (!existsSync(PDF)) {
  console.log('PDF 파일이 없어 건너뜁니다:', PDF);
  process.exit(0);
}
const pdfjs: any = await import('pdfjs-dist/legacy/build/pdf.mjs');
pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL('node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs').href;
const { parseTimetableItems } = await import('../src/timetablePdf');
const { representative } = await import('../src/daySheet');

let failed = 0;
const check = (cond: boolean, msg: string) => {
  if (!cond) {
    failed++;
    console.log('  FAIL:', msg);
  }
};

const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(PDF)), verbosity: 0 }).promise;
const pages = [];
for (let i = 1; i <= doc.numPages; i++) {
  const tc = await (await doc.getPage(i)).getTextContent();
  pages.push(tc.items.filter((it: any) => 'str' in it).map((it: any) => ({ str: it.str, x: it.transform[4], y: it.transform[5], w: it.width, h: it.height || Math.abs(it.transform[3]) })));
}
const tt = parseTimetableItems(pages);
check(!!tt, '시간표를 찾지 못함');
if (tt) {
  console.log('제목:', tt.title);
  check(tt.title === '2026학년도 2학기 중간고사', `제목 "${tt.title}"`);
  const want: Record<string, string> = { '2026-10-01': '1~3', '2026-10-02': '1~3', '2026-10-06': '1~4', '2026-10-07': '1~3', '2026-10-08': '1~3' };
  tt.days.forEach((d) => {
    console.log(`  ${d.date} ${d.start}~${d.end}교시`);
    check(want[d.date] === `${d.start}~${d.end}`, `${d.date} ${d.start}~${d.end} (기대 ${want[d.date]})`);
  });
  check(tt.days.length === 5, `날짜 수 ${tt.days.length}`);
  check(tt.exams.length === 30, `과목 수 ${tt.exams.length} (기대 30)`);
  const byKey = new Map(tt.exams.map((e) => [`${e.date.slice(5)}|${e.period}|${e.grade}`, e.text]));
  const expect: [string, string][] = [
    ['10-01|1|2학년', '화법과 언어'],
    ['10-01|2|2학년', '일본어 회화'],
    ['10-01|3|2학년', '인공지능 수학'],
    ['10-06|4|1학년', '한국사2'],
    ['10-08|3|3학년', '화학Ⅱ'],
    ['10-07|1|2학년', '미적분Ⅰ'],
  ];
  expect.forEach(([k, subject]) => check((byKey.get(k) ?? '').replace(/\s/g, '').startsWith(subject.replace(/\s/g, '')), `${k} → "${byKey.get(k)}" (기대 ${subject})`));
  console.log('과목', tt.exams.length, '건');
  ['10-01|1|2학년', '10-06|4|1학년', '10-01|2|3학년'].forEach((k) => console.log(`  ${k}: ${byKey.get(k)}  →  대표 ${representative(byKey.get(k) ?? '')}`));
}
console.log(failed ? `\n${failed} FAILED` : '\nALL CHECKS PASSED');
process.exit(failed ? 1 : 0);
