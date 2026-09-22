import { writeFileSync } from 'node:fs';

// 브라우저 전용 download()를 가로채 Blob을 잡아 둔다
const saved: Blob[] = [];
(globalThis as any).document = { createElement: () => ({ click() {}, remove() {} }), body: { appendChild() {} } };
(globalThis as any).URL.createObjectURL = (b: Blob) => {
  saved.push(b);
  return 'blob:x';
};
(globalThis as any).URL.revokeObjectURL = () => {};

const { downloadTemplate, importWorkbook } = await import('../src/excel');
const { demoProject } = await import('../src/demo');
const { cellKey, computeStats, emptyProject, getSlots, needKey, remapPreserve, validate } = await import('../src/model');

let failed = 0;
const check = (cond: boolean, msg: string) => {
  if (!cond) {
    failed++;
    console.log('  FAIL:', msg);
  }
};
const grab = async (name: string) => {
  const buf = Buffer.from(await saved.pop()!.arrayBuffer());
  try { writeFileSync(`/tmp/_tpl-${name}.xlsx`, buf); } catch {}
  return new File([buf], `${name}.xlsx`);
};

// 1) 빈 프로젝트 → 기본정보 시트만 있는 양식
{
  const p = emptyProject();
  await downloadTemplate(p, getSlots(p));
  const file = await grab('empty');
  const { project, notes } = await importWorkbook(file);
  console.log('\n[빈 양식] 시트 불러오기 notes:', notes[0]);
  check(project.days.length === 0 && project.rooms.length === 0, '빈 양식은 데이터가 없어야 함');
  check(project.teachers.length === 0, '빈 양식에 교사가 들어가면 안 됨(안내문이 데이터로 읽히는지 확인)');
}

// 2) 예시 데이터 → 3개 시트가 채워진 양식 → 다시 불러오기
{
  const p = demoProject();
  const slots = getSlots(p);
  p.teachers[0].target = 6;
  await downloadTemplate(p, slots);
  const file = await grab('full');
  const { project: back, notes } = await importWorkbook(file);
  console.log('\n[채운 양식] notes:', notes);
  const bs = getSlots(back);
  check(back.days.length === p.days.length, `일정 ${back.days.length} != ${p.days.length}`);
  check(back.rooms.length === p.rooms.length, `고사실 ${back.rooms.length} != ${p.rooms.length}`);
  check(back.teachers.length === p.teachers.length, `교사 ${back.teachers.length} != ${p.teachers.length}`);
  check(back.roles.length === p.roles.length, `보직 ${back.roles.length} != ${p.roles.length}`);
  check(bs.length === slots.length, `시험시간 ${bs.length} != ${slots.length}`);
  const st = computeStats(back, bs);
  const orig = computeStats(p, slots);
  check(st.totalNeed === orig.totalNeed, `필요 감독 수 ${st.totalNeed} != ${orig.totalNeed}`);
  const xCount = bs.reduce((a, s) => a + back.teachers.filter((t) => back.cells[cellKey(t.id, s.key)]?.x).length, 0);
  const fixedCount = bs.reduce((a, s) => a + back.teachers.filter((t) => back.cells[cellKey(t.id, s.key)]?.fixed).length, 0);
  const origX = slots.reduce((a, s) => a + p.teachers.filter((t) => p.cells[cellKey(t.id, s.key)]?.x).length, 0);
  const origFixed = slots.reduce((a, s) => a + p.teachers.filter((t) => p.cells[cellKey(t.id, s.key)]?.fixed).length, 0);
  check(xCount === origX, `감독 불가 ${xCount} != ${origX}`);
  check(fixedCount === origFixed, `고정 ${fixedCount} != ${origFixed}`);
  check(back.teachers[0].target === 6, `배정할 시간 ${back.teachers[0].target} != 6`);
  const forb = back.teachers.filter((t) => t.forbidden.length).length;
  check(forb === p.teachers.filter((t) => t.forbidden.length).length, `못 들어가는 고사실 교사 수 ${forb}`);
  check(back.teachers[3].prevLoad === p.teachers[3].prevLoad, '이전 누적 업무강도 보존');
  console.log('  일정', back.days.length, '고사실', back.rooms.length, '교사', back.teachers.length, '필요', st.totalNeed, 'x', xCount, '고정', fixedCount, '불가고사실교사', forb);
  const errs = validate(back, bs, st).filter((i) => i.level === 'error');
  check(errs.length === 0, '불러온 뒤 오류: ' + JSON.stringify(errs.slice(0, 3)));
}

// 3) 기본정보만 있는 양식을 다시 올릴 때 기존 필요 감독 수·표시가 유지되는지
{
  const prev = demoProject();
  const prevSlots = getSlots(prev);
  const basicOnly = { ...emptyProject(), days: prev.days, rooms: prev.rooms, roles: prev.roles, teachers: prev.teachers.map((t) => ({ ...t, forbidden: [], target: null })) };
  const merged = remapPreserve(prev, JSON.parse(JSON.stringify(basicOnly)));
  const ms = getSlots(merged);
  const mstat = computeStats(merged, ms);
  const pstat = computeStats(prev, prevSlots);
  console.log('\n[기본정보만 재업로드]');
  check(mstat.totalNeed === pstat.totalNeed, `필요 감독 수 유지 실패 ${mstat.totalNeed} != ${pstat.totalNeed}`);
  check(Object.keys(merged.cells).length === Object.keys(prev.cells).length, '감독 불가·고정 표시 유지 실패');
  check(merged.teachers.some((t) => t.forbidden.length > 0), '못 들어가는 고사실 유지 실패');
  const needSame = ms.every((s) => merged.roles.every((r) => merged.rooms.every((m) => (merged.need[needKey(s.key, r.id, m.id)] || 0) === (prev.need[needKey(s.key, r.id, m.id)] || 0))));
  check(needSame, '필요 감독 수 값이 원본과 다름');
  console.log('  필요', mstat.totalNeed, '칸', Object.keys(merged.cells).length, '불가고사실교사', merged.teachers.filter((t) => t.forbidden.length).length);
}

console.log(failed ? `\n${failed} FAILED` : '\nALL CHECKS PASSED');
process.exit(failed ? 1 : 0);
