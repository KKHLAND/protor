import { readFileSync, writeFileSync } from 'node:fs';
import ExcelJS from 'exceljs';

// 브라우저 전용 download()를 가로채 파일로 저장
const saved: { name: string; blob: Blob }[] = [];
(globalThis as any).document = {
  createElement: () => ({ click() {}, remove() {} }),
  body: { appendChild() {} },
};
(globalThis as any).URL.createObjectURL = (b: Blob) => {
  saved.push({ name: '', blob: b });
  return 'blob:x';
};
(globalThis as any).URL.revokeObjectURL = () => {};

const { importWorkbook, exportWorkbook } = await import('../src/excel');
const { demoProject } = await import('../src/demo');
const { autoFillTargets, buildEngineInput, applyGrid, computeStats, getSlots, validate } = await import('../src/model');
const { runBest } = await import('../src/engine');

// 1) 원본 proctor.xlsm 불러오기
const buf = readFileSync('../proctor.xlsm');
const file = new File([buf], 'proctor.xlsm');
const { project, notes } = await importWorkbook(file);
console.log('[xlsm] notes', notes);
console.log('[xlsm] days', project.days, 'rooms', project.rooms.map((r) => r.name), 'teachers', project.teachers.map((t) => `${t.name}:${t.prevLoad}`), 'roles', project.roles);

// 2) 예시 데이터 배정 → 내보내기 → 다시 불러오기(내장 데이터 & 시트 파싱 둘 다)
let p = demoProject();
const slots = getSlots(p);
p = autoFillTargets(p, slots).project;
const res = runBest(buildEngineInput(p, slots), 5, 7);
if (res.status !== 'ok') throw new Error('engine failed');
p = applyGrid(p, slots, res.grid);
await exportWorkbook(p, slots, computeStats(p, slots));
const out = Buffer.from(await saved[0].blob.arrayBuffer());
writeFileSync('scripts/_export-test.xlsx', out);
console.log('[export] bytes', out.length);

const back = await importWorkbook(new File([out], 'x.xlsx'));
console.log('[roundtrip _data] notes', back.notes, 'same cells', JSON.stringify(back.project.cells) === JSON.stringify(p.cells));

// 내장 데이터 시트를 지우고 시트 파싱 경로 검증
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(out);
wb.removeWorksheet(wb.getWorksheet('_data')!.id);
const stripped = Buffer.from(await wb.xlsx.writeBuffer());
const back2 = await importWorkbook(new File([stripped], 'y.xlsx'));
const s2 = getSlots(back2.project);
const st2 = computeStats(back2.project, s2);
const errs = validate(back2.project, s2, st2).filter((i) => i.level === 'error');
console.log('[roundtrip sheets] notes', back2.notes);
console.log('[roundtrip sheets] need total', st2.totalNeed, 'hasResults', st2.hasResults, 'errors', errs.length, errs.slice(0, 3));
