// 강사처럼 특정 날짜에 정해진 시수만 감독하는 교사 처리 검증
const saved: Blob[] = [];
(globalThis as any).document = { createElement: () => ({ click() {}, remove() {} }), body: { appendChild() {} } };
(globalThis as any).URL.createObjectURL = (b: Blob) => {
  saved.push(b);
  return 'blob:x';
};
(globalThis as any).URL.revokeObjectURL = () => {};

const { demoProject } = await import('../src/demo');
const { applyGrid, autoFillTargets, buildEngineInput, cellKey, computeStats, getSlots, validate } = await import('../src/model');
const { runBest, runEngine } = await import('../src/engine');
const { downloadTemplate, importWorkbook } = await import('../src/excel');

let failed = 0;
const check = (cond: boolean, msg: string) => {
  if (!cond) {
    failed++;
    console.log('  FAIL:', msg);
  }
};

// 강사 2명 설정: 10/13 하루만 2시간, 10/14~15 이틀만 3시간
let p = demoProject();
const slots = getSlots(p);
const dayIds = p.days.map((d) => d.id);
p.teachers[10] = { ...p.teachers[10], availableDays: [dayIds[1]], target: 2, lockTarget: true };
p.teachers[11] = { ...p.teachers[11], availableDays: [dayIds[2], dayIds[3]], target: 3, lockTarget: true };
const lecturerNames = [p.teachers[10].name, p.teachers[11].name];

// 1) 자동 채우기가 고정 시수를 건드리지 않는지
const filled = autoFillTargets(p, slots);
p = filled.project;
console.log('\n[자동 채우기]', filled.message ?? '정상');
check(p.teachers[10].target === 2, `강사1 시수 ${p.teachers[10].target} != 2`);
check(p.teachers[11].target === 3, `강사2 시수 ${p.teachers[11].target} != 3`);
const totalTarget = p.teachers.reduce((a, t) => a + (t.target || 0), 0);
const stats0 = computeStats(p, slots);
check(totalTarget === stats0.totalNeed, `시간 합 ${totalTarget} != 필요 ${stats0.totalNeed}`);
console.log('  강사 시수', p.teachers[10].target, p.teachers[11].target, '/ 전체 합', totalTarget, '= 필요', stats0.totalNeed);
check(stats0.teachers[10].maxPossible === 3, `강사1 감독 가능 시간 ${stats0.teachers[10].maxPossible} != 3 (10/13 3교시)`);

// 2) 배정 결과가 조건을 지키는지
const res = runBest(buildEngineInput(p, slots), 10, 11);
console.log('[배정]', res.status);
if (res.status !== 'ok') {
  console.log(res);
  failed++;
} else {
  const out = applyGrid(p, slots, res.grid);
  const st = computeStats(out, slots);
  const errs = validate(out, slots, st).filter((i) => i.level === 'error');
  check(errs.length === 0, '검증 오류: ' + JSON.stringify(errs.slice(0, 3)));
  [10, 11].forEach((ti, k) => {
    const t = out.teachers[ti];
    const mine = slots.filter((s) => {
      const c = out.cells[cellKey(t.id, s.key)];
      return c?.on || c?.fixed;
    });
    check(mine.length === t.target, `${t.name} 배정 ${mine.length} != 시수 ${t.target}`);
    const okDays = mine.every((s) => t.availableDays!.includes(s.dayId));
    check(okDays, `${t.name} 감독 가능 날짜 밖에 배정됨`);
    console.log(`  ${t.name}: ${mine.length}시간 · ${[...new Set(mine.map((s) => s.date))].join(', ')} · 날짜조건 ${okDays ? 'OK' : 'NG'}`);
    void k;
  });
  check(st.totalNeed === 297, '필요 시간 변동');
  console.log('  충돌', res.metrics.conflicts, '3연속', res.metrics.consec3, '업무강도', res.metrics.loadMin, '~', res.metrics.loadMax);
}

// 3) 가능 날짜보다 시수가 많으면 미리 걸러지는지
{
  const bad = { ...p, teachers: p.teachers.map((t, i) => (i === 10 ? { ...t, target: 5 } : t)) };
  const r2 = runEngine(buildEngineInput(bad, slots), { seed: 1 });
  check(r2.status === 'precheck', `초과 시수인데 ${r2.status}`);
  if (r2.status === 'precheck') console.log('\n[초과 시수 점검]', r2.issues.find((i) => i.includes(lecturerNames[0])) ?? r2.issues[0]);
}

// 4) 엑셀 양식 왕복에서 조건이 유지되는지
{
  await downloadTemplate(p, slots);
  const buf = Buffer.from(await saved.pop()!.arrayBuffer());
  const { project: back, notes } = await importWorkbook(new File([buf], 'tpl.xlsx'));
  const b10 = back.teachers.find((t) => t.name === lecturerNames[0])!;
  const b11 = back.teachers.find((t) => t.name === lecturerNames[1])!;
  const bs = getSlots(back);
  const dayOf = (id: string) => back.days.find((d) => d.id === id)?.date;
  console.log('\n[엑셀 왕복]', notes.find((n) => n.includes('근무 조건')) ?? notes[2]);
  check(!!b10.lockTarget && b10.target === 2, `강사1 시수 고정 유지 실패: ${b10.target} ${b10.lockTarget}`);
  check(b10.availableDays?.length === 1 && dayOf(b10.availableDays[0]) === '2026-10-13', `강사1 가능 날짜 유지 실패: ${b10.availableDays?.map(dayOf)}`);
  check(b11.availableDays?.length === 2, `강사2 가능 날짜 유지 실패: ${b11.availableDays?.map(dayOf)}`);
  const bst = computeStats(back, bs);
  check(bst.teachers[back.teachers.indexOf(b10)].maxPossible === 3, '불러온 뒤 감독 가능 시간 계산 오류');
  console.log('  강사1', b10.target, '시간 고정', !!b10.lockTarget, '가능일', b10.availableDays?.map(dayOf).join(','), '/ 강사2 가능일', b11.availableDays?.map(dayOf).join(','));
}

console.log(failed ? `\n${failed} FAILED` : '\nALL CHECKS PASSED');
process.exit(failed ? 1 : 0);
