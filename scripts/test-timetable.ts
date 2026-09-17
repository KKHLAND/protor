// 시험 시간표(엑셀) 읽기 → 일정 적용 → 날짜별 시감표 생성 → 엑셀 변환 검증
import ExcelJS from 'exceljs';

(globalThis as any).document = { createElement: () => ({ click() {}, remove() {} }), body: { appendChild() {} } };

const { parseTimetable, applyTimetable } = await import('../src/timetable');
const { buildDayBook, representative } = await import('../src/daySheet');
const { sheetBookToXlsx } = await import('../src/excel');
const { appsScriptCode } = await import('../src/appsScript');
const { demoProject } = await import('../src/demo');
const { applyGrid, autoFillTargets, buildEngineInput, cellKey, getSlots } = await import('../src/model');
const { runBest } = await import('../src/engine');

let failed = 0;
const check = (cond: boolean, msg: string) => {
  if (!cond) {
    failed++;
    console.log('  FAIL:', msg);
  }
};

// 실제 "2026학년도 2학기 중간고사 시간표 및 시험범위" 시트와 같은 구조의 파일
function fixture(dateCells: 'text' | 'date') {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('시간표');
  ws.getCell('B2').value = '2026학년도 2학기 중간고사 시간표';
  ws.mergeCells('B2:M2');
  const dates = [1, 2, 6, 7, 8];
  const labels = ['10월 1일(목)', '10월 2일(금)', '10월 6일(화)', '10월 7일(수)', '10월 8일(목)'];
  const header = (r: number, oneGrade: boolean) => {
    ws.getCell(r, 2).value = '교시';
    ws.mergeCells(r, 2, r + 1, 2);
    ws.getCell(r, 3).value = '본령';
    ws.mergeCells(r, 3, r + 1, 3);
    dates.forEach((d, i) => {
      const c = 4 + 2 * i;
      ws.getCell(r, c).value = dateCells === 'date' ? new Date(Date.UTC(2026, 9, d)) : labels[i];
      ws.mergeCells(r, c, r, c + 1);
      if (oneGrade) {
        ws.getCell(r + 1, c).value = '1학년';
        ws.mergeCells(r + 1, c, r + 1, c + 1);
      } else {
        ws.getCell(r + 1, c).value = '2학년';
        ws.getCell(r + 1, c + 1).value = '3학년';
      }
    });
  };
  header(4, false);
  const rowsA: [number, string, string[]][] = [
    [1, '8:20', ['화법과 언어 (40분/공통) (8:20~9:00) 교사01, 교사02', '영어 독해와 작문 (40분/A~H) (8:20~9:00) 교사03, 교사04', '영어Ⅱ (40분/공통) (8:20~9:00) 교사05, 교사06', '독서 (40분/G,H,I,J) (8:20~9:00) 교사07', '사회와 문화 (40분/A,B,C,D,E) (8:20~9:00) 교사08', '확률과 통계 (40분/B,C,D,E,F,,I,J) (8:20~9:00) 교사09, 교사10', '미적분Ⅰ (50분/공통) (8:20~9:10) 교사11, 교사12, 교사10, 교사13', '미적분 (50분/D,E,F) (8:20~9:10) 교사11', '윤리와 사상 (40분/A,B,C,D) (8:20~9:00) 교사14', '사회문제 탐구 (30분/A,B,C,E,F,J) ※8:20~8:30 대기 (8:30~9:00) 교사15, 교사16']],
    [2, '9:30', ['일본어 회화 (40분/A,B,C,D) (9:30~10:10) 교사17, 교사18', '물리학Ⅱ(B,F) / 정치와 법(A) (40분/9:30~10:10) 교사09 / 교사16', '화학 반응의 세계(서) (40분/A,B,D,E) (9:30~10:10) 교사19', '지구과학Ⅱ(G) / 세계지리(H) (40분/9:30~10:10) 교사20 / 교사21', '동아시아 역사 기행(B,C) / 역학과 에너지(서)(A,D,E) (40분/9:30~10:10) 교사22 / 교사23', '생활과 과학 (30분/A,D,G,H,I) ※9:30~9:40 대기 (9:40~10:10) 교사09, 교사24', '세계 문화와 영어(서) (40분/A,D,E) (9:30~10:10) 교사04, 교사05', '여행지리 (30분/E,J) ※9:30~9:40 대기 (9:40~10:10) 교사25', '행성우주과학 (40분/C) (9:30~10:10) 교사23', '화법과 작문 (40분/C,G,H,I,J) (9:30~10:10) 교사26']],
    [3, '10:40', ['인공지능 수학 (30분/A,B,D) ※10:40~10:50 대기 (10:50~11:20) 교사27', '', '', '', '', '윤리와 사상 (40분/A,D,G,H) (10:40~11:20) 교사28', '세포와 물질대사 (40분/A,C,E) ※10:40~10:50 대기 (10:50~11:30) 교사24', '', '', '화학Ⅱ (30분/I,J) ※10:40~10:50 대기 (10:50~11:20) 교사29']],
  ];
  rowsA.forEach(([per, time, texts], k) => {
    const r = 6 + k;
    ws.getCell(r, 2).value = per;
    ws.getCell(r, 3).value = time;
    texts.forEach((t, i) => (ws.getCell(r, 4 + i).value = t || null));
  });
  header(10, true);
  const rowsB: [number, string, string[]][] = [
    [3, '10:40', ['통합과학2 (40분) (10:40~11:20) 교사30, 교사31, 교사23(물리), 교사32', '공통영어2(서) (40분) (10:40~11:20) 교사33, 교사34, 교사35', '공통국어2 (40분) (10:40~11:20) 교사36, 교사37', '공통수학2(서) (50분) (10:40~11:30) 교사38, 교사39', '통합사회2 (40분) (10:40~11:20) 교사15, 교사14, 교사28, 교사25']],
    [4, '11:40', ['', '', '한국사2 (40분) (11:40~12:20) 교사22, 교사40', '', '']],
  ];
  rowsB.forEach(([per, time, texts], k) => {
    const r = 12 + k;
    ws.getCell(r, 2).value = per;
    ws.getCell(r, 3).value = time;
    texts.forEach((t, i) => {
      const c = 4 + 2 * i;
      ws.getCell(r, c).value = t || null;
      ws.mergeCells(r, c, r, c + 1);
    });
  });
  ws.getCell('B15').value = '※ 응시 유의사항';
  ws.getCell('B20').value = '3. 시험 장소 변경';
  ws.getCell('B21').value = '학년';
  ws.getCell('D21').value = '시험 일시';
  ws.getCell('B22').value = 3;
  ws.getCell('D22').value = '10월 6일(화) 1교시';
  ws.mergeCells('D22:E22');
  ws.getCell('K22').value = '확률과 통계J';
  ws.getCell('B24').value = '4. 시험 없는 시간 대기 장소';
  ws.getCell('B25').value = '구분';
  ws.mergeCells('B25:C26');
  labels.forEach((l, i) => {
    ws.getCell(25, 4 + 2 * i).value = l;
    ws.mergeCells(25, 4 + 2 * i, 25, 5 + 2 * i);
    ws.getCell(26, 4 + 2 * i).value = '2학년';
    ws.getCell(26, 5 + 2 * i).value = '3학년';
  });
  ws.getCell('B27').value = '1교시 대기실';
  ws.mergeCells('B27:C27');
  ['X', 'X', 'X', '대기실B (22명)', '대기실A/2-8반 (60명)', '대기실B (33명)', 'X', '대기실B/대기실C (38명)', '대기실A (17명)', '대기실B/대기실C (37명)'].forEach(
    (v, i) => (ws.getCell(27, 4 + i).value = v),
  );
  return wb;
}

const expected: Record<string, [number, number]> = {
  '2026-10-01': [1, 3],
  '2026-10-02': [1, 3],
  '2026-10-06': [1, 4],
  '2026-10-07': [1, 3],
  '2026-10-08': [1, 3],
};

for (const kind of ['text', 'date'] as const) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await fixture(kind).xlsx.writeBuffer());
  const tt = parseTimetable(wb);
  console.log(`\n[시간표 읽기 · 날짜 칸=${kind}]`, tt ? `${tt.title} / ${tt.days.length}일 / 과목 ${tt.exams.length}건` : 'null');
  check(!!tt, '시간표를 찾지 못함');
  if (!tt) continue;
  check(tt.title === '2026학년도 2학기 중간고사', `제목 ${tt.title}`);
  check(tt.days.length === 5, `날짜 수 ${tt.days.length}`);
  tt.days.forEach((d) => {
    const e = expected[d.date];
    check(!!e && e[0] === d.start && e[1] === d.end, `${d.date} ${d.start}~${d.end}교시 (기대 ${e})`);
  });
  check(tt.exams.length === 30, `과목 수 ${tt.exams.length} (기대 30)`);
  console.log('  ', tt.days.map((d) => `${d.date.slice(5)} ${d.start}~${d.end}`).join(' | '));
}

// 대표 교사 표기
console.log('\n[대표 교사]');
[
  ['화법과 언어 (40분/공통) (8:20~9:00) 교사01, 교사02', '화법과 언어 교사01'],
  ['물리학Ⅱ(B,F) / 정치와 법(A) (40분/9:30~10:10) 교사09 / 교사16', '물리학Ⅱ 교사09, 정치와 법 교사16'],
  ['통합과학2 (40분) (10:40~11:20) 교사30, 교사31, 교사23(물리), 교사32', '통합과학2 교사30'],
  ['사회문제 탐구 (30분/A,B,C,E,F,J) ※8:20~8:30 대기 (8:30~9:00) 교사15, 교사16', '사회문제 탐구 교사15'],
].forEach(([src, want]) => {
  const got = representative(src);
  check(got === want, `대표 교사 "${got}" != "${want}"`);
  console.log('  ', got);
});

// 예시 데이터에 시간표를 적용하고 배정 → 1일차 시감표
{
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await fixture('text').xlsx.writeBuffer());
  const tt = parseTimetable(wb)!;
  let p = demoProject();
  const applied = applyTimetable(p, tt);
  p = applied.project;
  console.log('\n[일정 적용]', applied.notes.slice(0, 3).join(' / '));
  const slots = getSlots(p);
  check(p.days.length === 5 && slots.length === 16, `적용 후 일정 ${p.days.length}일 · ${slots.length}시간`);
  check(Object.keys(p.exams ?? {}).length === 30, `저장된 과목 ${Object.keys(p.exams ?? {}).length}`);
  // 새 날짜에 필요 감독 수 채우기 (예시 규칙 그대로)
  for (const s of slots) {
    p.rooms.filter((r) => r.id !== 'rwait').forEach((r) => (p.need[`${s.key}|main|${r.id}`] = 1));
    p.rooms.filter((r) => r.name.startsWith('3-')).forEach((r) => (p.need[`${s.key}|sub|${r.id}`] = 1));
    p.need[`${s.key}|wait|rwait`] = 1;
  }
  p = autoFillTargets(p, slots).project;
  const res = runBest(buildEngineInput(p, slots), 3, 9);
  check(res.status === 'ok', `배정 ${res.status}`);
  if (res.status === 'ok') {
    p = applyGrid(p, slots, res.grid);
    const dayIdx = 2; // 10/6 (1~4교시)
    const book = buildDayBook(p, slots, dayIdx);
    console.log('\n[시감표]', book.title, '/ 시트:', book.sheets.map((s) => s.name).join(', '));
    check(book.title === '2026학년도 2학기 중간고사 시감표_3일차', `제목 ${book.title}`);
    check(book.sheets.length === 4, `시트 수 ${book.sheets.length}`);
    const s1 = book.sheets[0];
    check(String(s1.rows[0][0]?.v) === '2026학년도 2학기 중간고사 3일차(10월 6일)', `머리말 ${s1.rows[0][0]?.v}`);
    const flat = (spec: typeof s1) => spec.rows.flat().map((c) => (c?.v ?? '').toString());
    // 그날 배정 수 = 시감표 이름 칸 수
    const daySlots = slots.filter((s) => s.dayIdx === dayIdx);
    let assigned = 0;
    let waiting = 0;
    p.teachers.forEach((t) =>
      daySlots.forEach((s) => {
        const c = p.cells[cellKey(t.id, s.key)];
        if ((c?.on || c?.fixed) && c.role) {
          assigned++;
          if (c.room === 'rwait') waiting++;
        }
      }),
    );
    const roomStart = s1.rows.findIndex((r) => r[0]?.v === '학급') + 1;
    const nameCount = s1.rows
      .slice(roomStart, roomStart + p.rooms.length)
      .flatMap((r) => r.slice(1))
      .reduce((a, c) => a + (c?.v ? String(c.v).split('\n').length : 0), 0);
    check(nameCount === assigned, `시감표 이름 ${nameCount} != 배정 ${assigned}`);
    const waitRow = s1.rows.find((r) => r[0]?.v === '대기실')!;
    const waitInMain = waitRow.slice(1).filter((_, i) => i % 2 === 0).filter((c) => c?.v).length;
    const waitInSub = waitRow.slice(1).filter((_, i) => i % 2 === 1).filter((c) => c?.v).length;
    check(waitInMain === waiting && waitInSub === 0, `대기실 감독이 정감독 칸에 있어야 함 (정 ${waitInMain}, 부 ${waitInSub})`);
    const gradeRows = s1.rows.filter((r) => /학년$/.test(String(r[0]?.v ?? ''))).map((r) => r[0]!.v);
    check(gradeRows.join(',') === '1학년,2학년,3학년', `학년 줄 ${gradeRows}`);
    const rep = s1.rows.find((r) => r[0]?.v === '대표 교사');
    check(!!rep && String(rep[1]?.v).includes('2학년 사회와 문화 교사08'), `대표 교사 줄 ${rep?.[1]?.v}`);
    check(!!rep && String(rep[7]?.v).includes('1학년 한국사2 교사22'), `4교시 대표 교사 ${rep?.[7]?.v}`);
    const s2 = book.sheets[1];
    check(flat(s2).includes(`${new Set(p.teachers.filter((t) => daySlots.some((s) => p.cells[cellKey(t.id, s.key)]?.role)).map((t) => t.id)).size} / ${p.teachers.length}`), '교사별 합계 칸');
    const s3 = book.sheets[2];
    const todaySum = s3.rows.slice(2).reduce((a, r) => a + (Number(r[3]?.v) || 0), 0);
    check(todaySum === assigned, `감독 누계 시트 소계 합 ${todaySum} != ${assigned}`);
    console.log(`   배정 ${assigned}건 (대기실 ${waiting}) · 시감표 이름 ${nameCount} · 소계 합 ${todaySum}`);

    // 엑셀 변환
    const blob = await sheetBookToXlsx(book);
    const back = new ExcelJS.Workbook();
    await back.xlsx.load(Buffer.from(await blob.arrayBuffer()));
    const w1 = back.getWorksheet('시감표')!;
    check(w1.getCell(1, 1).value === '2026학년도 2학기 중간고사 3일차(10월 6일)', '엑셀 머리말');
    check(((w1.model as any).merges?.length ?? 0) > 5, '엑셀 병합 셀');
    check(back.worksheets.map((w) => w.name).join(',') === '시감표,교사별,감독 누계,시험 시간표', '엑셀 시트 이름');
    console.log('   엑셀', blob.size, 'bytes · 병합', (w1.model as any).merges?.length);
  }
}

// 앱스 스크립트 코드 문법
try {
  new Function(appsScriptCode('abc123'));
  console.log('\n[앱스 스크립트] 문법 OK');
} catch (e) {
  check(false, `앱스 스크립트 문법 오류: ${(e as Error).message}`);
}

console.log(failed ? `\n${failed} FAILED` : '\nALL CHECKS PASSED');
process.exit(failed ? 1 : 0);
