import type { Cell, Project } from './types';
import { cellKey, getSlots, needKey } from './model';

const FAMILY = ['김', '이', '박', '최', '정', '강', '조', '윤', '장', '임', '한', '오', '서', '신', '권', '황', '안', '송', '류', '홍'];
const GIVEN = ['민준', '서연', '도윤', '지우', '하준', '서윤', '은우', '지민', '시우', '하윤', '준서', '수아', '예준', '채원', '유준', '지아', '건우', '다은', '현우', '소율', '우진', '예린', '선우', '윤서', '연우'];

/** 예시 데이터: 4일 · 18개 고사실 · 52명 · 보직 3종 */
export function demoProject(): Project {
  const days = [
    { id: 'd1', date: '2026-10-12', start: 1, end: 3 },
    { id: 'd2', date: '2026-10-13', start: 1, end: 3 },
    { id: 'd3', date: '2026-10-14', start: 1, end: 3 },
    { id: 'd4', date: '2026-10-15', start: 1, end: 2 },
  ];
  const examRooms = [1, 2, 3].flatMap((g) => [1, 2, 3, 4, 5, 6].map((c) => ({ id: `r${g}${c}`, name: `${g}-${c}` })));
  // 고교학점제로 시험이 없는 학생들이 머무는 대기실
  const rooms = [...examRooms, { id: 'rwait', name: '대기실' }];
  const roles = [
    { id: 'main', name: '정감독', weight: 100 },
    { id: 'sub', name: '부감독', weight: 50 },
    { id: 'hall', name: '복도감독', weight: 30 },
    { id: 'wait', name: '대기실 감독', weight: 30 },
  ];
  const teachers = Array.from({ length: 52 }, (_, i) => ({
    id: `t${i + 1}`,
    name: FAMILY[(i * 7) % FAMILY.length] + GIVEN[(i * 11) % GIVEN.length],
    prevLoad: [0, 50, 100, 150, 200][(i * 3) % 5],
    target: null as number | null,
    forbidden: [] as string[],
  }));
  // 담임 교사는 자기 반 고사실에 들어가지 않는다
  rooms.forEach((r, i) => {
    teachers[i].forbidden = [r.id];
  });

  const p: Project = {
    version: 1,
    meta: { school: 'OO고등학교', title: '2026학년도 2학기 1회고사' },
    days,
    rooms,
    roles,
    teachers,
    need: {},
    cells: {},
    extra: {},
  };
  const slots = getSlots(p);
  for (const s of slots) {
    for (const r of examRooms) p.need[needKey(s.key, 'main', r.id)] = 1;
    for (const r of examRooms.filter((r) => r.name.startsWith('3-'))) p.need[needKey(s.key, 'sub', r.id)] = 1;
    for (const r of ['r11', 'r21', 'r31']) p.need[needKey(s.key, 'hall', r)] = 1;
    // 대기실은 1교시에 대기 학생이 많아 2명, 나머지 교시는 1명
    p.need[needKey(s.key, 'wait', 'rwait')] = s.period === 1 ? 2 : 1;
  }
  const cells: Record<string, Cell> = {};
  const mark = (t: number, day: string, periods: number[], c: Cell) =>
    periods.forEach((per) => (cells[cellKey(`t${t}`, `${day}:${per}`)] = c));
  mark(40, 'd1', [1, 2, 3], { x: true }); // 출장
  mark(41, 'd2', [1, 2, 3], { x: true });
  mark(42, 'd3', [2, 3], { x: true });
  mark(43, 'd4', [1, 2], { x: true });
  mark(44, 'd1', [1], { x: true });
  mark(45, 'd2', [3], { x: true });
  mark(50, 'd1', [1], { fixed: true, fixedRole: 'hall' });
  mark(51, 'd3', [1], { fixed: true });
  p.cells = cells;
  return p;
}
