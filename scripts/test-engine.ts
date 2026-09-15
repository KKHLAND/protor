import { demoProject } from '../src/demo';
import { applyGrid, autoFillTargets, buildEngineInput, cellKey, computeStats, getSlots, validate } from '../src/model';
import { runBest, runEngine } from '../src/engine';

let failed = 0;
const check = (cond: boolean, msg: string) => {
  if (!cond) {
    failed++;
    console.log('  FAIL:', msg);
  }
};

function verify(label: string, trials: number, seed: number) {
  let p = demoProject();
  const slots = getSlots(p);
  p = autoFillTargets(p, slots).project;
  const inp = buildEngineInput(p, slots);
  const t0 = performance.now();
  const res = runBest(inp, trials, seed);
  const ms = performance.now() - t0;
  console.log(`\n[${label}] status=${res.status} ${ms.toFixed(0)}ms`);
  if (res.status !== 'ok') {
    console.log(res);
    failed++;
    return;
  }
  console.log('  metrics', res.metrics, 'warnings', res.warnings);
  const out = applyGrid(p, slots, res.grid);
  const stats = computeStats(out, slots);
  const issues = validate(out, slots, stats).filter((i) => i.level === 'error');
  check(issues.length === 0, 'validate errors: ' + JSON.stringify(issues.slice(0, 5)));
  out.teachers.forEach((t, ti) => {
    check(stats.teachers[ti].total === t.target, `${t.name} total ${stats.teachers[ti].total} != target ${t.target}`);
    slots.forEach((s) => {
      const c = out.cells[cellKey(t.id, s.key)];
      if (c?.x) check(!c.on && !c.role, `${t.name} ${s.key} assigned on x`);
      if (c?.fixed) check(!!c.role && !!c.room, `${t.name} ${s.key} fixed without assignment`);
      if (c?.fixed && c.fixedRole) check(c.role === c.fixedRole, `${t.name} fixed role mismatch`);
    });
  });
  const loads = stats.teachers.map((s) => s.cumLoad);
  console.log('  cumLoad range', Math.min(...loads), Math.max(...loads), 'maxConsec', Math.max(...stats.teachers.map((s) => s.maxConsec)));
  const dayImb = stats.teachers.map((s) => Math.max(...s.dayCounts) - Math.min(...s.dayCounts));
  console.log('  day spread histogram', dayImb.reduce((a: Record<number, number>, v) => ((a[v] = (a[v] || 0) + 1), a), {}));
}

verify('single', 1, 1);
verify('best of 10', 10, 42);

// 가용 인원이 부족한 교시가 생기면 needExtra를 돌려주는지
{
  let p = demoProject();
  const slots = getSlots(p);
  p = autoFillTargets(p, slots).project;
  // 첫 교시에 거의 모두 불가 처리 → 남은 교사들의 목표시간이 이미 다른 곳에 묶이도록
  const s0 = slots[0].key;
  p.teachers.slice(0, 24).forEach((t) => (p.cells[cellKey(t.id, s0)] = { x: true }));
  // 나머지 교사 목표시간을 조정: 남은 28명 중 27명이 필요하지만 한 명은 목표 1시간뿐이고 이미 고정
  const inp = buildEngineInput(p, slots);
  const res = runEngine(inp, { seed: 3 });
  console.log('\n[tight] status', res.status, res.status === 'precheck' ? res.issues.slice(0, 3) : res.status === 'needExtra' ? { slot: res.slot, n: res.candidates.length } : '');
}

// 교시별 인원은 충분해 보여도 전체 구조상 불가능한 경우 → needExtra
{
  const F = false;
  const inp = {
    T: 4, S: 3, R: 1, M: 2,
    slotDay: [0, 0, 0],
    roleWeight: [100],
    need: [[[1, 1]], [[1, 1]], [[0, 0]]],
    x: [[F, F, true], [F, F, true], [true, true, F], [true, true, F]],
    fixed: [[F, F, F], [F, F, F], [F, F, F], [F, F, F]],
    fixedRole: [[-1, -1, -1], [-1, -1, -1], [-1, -1, -1], [-1, -1, -1]],
    target: [1, 1, 1, 1],
    extra: [0, 0, 0, 0],
    forbidden: [[F, F], [F, F], [F, F], [F, F]],
    prevLoad: [0, 0, 0, 0],
    names: { teacher: ['A', 'B', 'C', 'D'], slot: ['s1', 's2', 's3'], role: ['정'], room: ['1', '2'] },
  };
  const res = runEngine(inp, { seed: 1 });
  check(res.status === 'needExtra', 'needExtra expected, got ' + res.status);
  if (res.status === 'needExtra') {
    console.log('\n[needExtra] slot', res.slot, 'candidates', res.candidates);
    inp.extra = [1, 1, 0, 0];
    const res2 = runEngine(inp, { seed: 1 });
    check(res2.status === 'ok', 'after extra should be ok, got ' + res2.status);
  }
}

// 합계 불일치 사전 점검
{
  const p = demoProject();
  const slots = getSlots(p);
  const res = runEngine(buildEngineInput(p, slots), { seed: 1 });
  check(res.status === 'precheck', 'precheck expected when targets empty');
  console.log('\n[precheck] ', res.status === 'precheck' ? res.issues[0] : res.status);
}

console.log(failed ? `\n${failed} FAILED` : '\nALL CHECKS PASSED');
process.exit(failed ? 1 : 0);
