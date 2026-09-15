import type { Cell, Day, Project, Slot } from './types';
import type { EngineInput, GridCell } from './engine';

export const uid = () => Math.random().toString(36).slice(2, 10);

export const ROLE_COLORS = ['#CFDDFB', '#E0D8FB', '#CDEFE2', '#FDE0D2', '#FFF0B8', '#D5ECFA', '#EAD7F3', '#E3E6EE'];
export const roleColor = (i: number) => (i < 0 ? 'transparent' : ROLE_COLORS[i % ROLE_COLORS.length]);

export const docTitle = (p: Project) => [p.meta.school, p.meta.title].map((s) => s.trim()).filter(Boolean).join(' ');

export const needKey = (slotKey: string, roleId: string, roomId: string) => `${slotKey}|${roleId}|${roomId}`;
export const cellKey = (teacherId: string, slotKey: string) => `${teacherId}|${slotKey}`;

export function emptyProject(): Project {
  return {
    version: 1,
    meta: { school: '', title: '' },
    days: [],
    rooms: [],
    roles: [
      { id: uid(), name: '정감독', weight: 100 },
      { id: uid(), name: '부감독', weight: 50 },
    ],
    teachers: [],
    need: {},
    cells: {},
    extra: {},
  };
}

const WEEK = ['일', '월', '화', '수', '목', '금', '토'];

export function parseDate(date: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
export function fmtMD(date: string) {
  const d = parseDate(date);
  if (!d) return '날짜 미정';
  return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}
export function fmtDay(date: string) {
  const d = parseDate(date);
  if (!d) return '날짜 미정';
  return `${fmtMD(date)}(${WEEK[d.getDay()]})`;
}
export const slotLabel = (s: Slot) => `${fmtDay(s.date)} ${s.period}교시`;

export function sortedDays(p: Project): Day[] {
  return p.days
    .map((d, i) => ({ d, i }))
    .sort((a, b) => {
      if (!a.d.date && !b.d.date) return a.i - b.i;
      if (!a.d.date) return 1;
      if (!b.d.date) return -1;
      return a.d.date.localeCompare(b.d.date) || a.i - b.i;
    })
    .map((x) => x.d);
}

export function getSlots(p: Project): Slot[] {
  const slots: Slot[] = [];
  sortedDays(p).forEach((d, dayIdx) => {
    const start = Math.max(1, Math.floor(d.start || 1));
    const end = Math.max(start, Math.floor(d.end || start));
    for (let period = start; period <= end; period++) {
      slots.push({
        key: `${d.id}:${period}`,
        dayId: d.id,
        dayIdx,
        date: d.date,
        period,
        posInDay: period - start,
        firstOfDay: period === start,
        lastOfDay: period === end,
      });
    }
  });
  return slots;
}

export const isAssigned = (c?: Cell) => !!c && (c.fixed || c.on);

export interface TeacherStat {
  total: number;
  fixedCount: number;
  xCount: number;
  maxPossible: number;
  roleCounts: number[];
  curLoad: number;
  cumLoad: number;
  dayCounts: number[];
  maxConsec: number;
  conflicts: number;
}

export interface SlotStat {
  need: number;
  available: number;
  assigned: number;
  fixed: number;
  needByRole: number[];
  assignedByRole: number[];
}

export interface Stats {
  teachers: TeacherStat[];
  slots: SlotStat[];
  totalNeed: number;
  totalTarget: number;
  hasResults: boolean;
  dayCount: number;
}

export function computeStats(p: Project, slots: Slot[]): Stats {
  const roleIdx = new Map(p.roles.map((r, i) => [r.id, i]));
  const dayCount = slots.length ? slots[slots.length - 1].dayIdx + 1 : 0;
  let hasResults = false;
  const slotStats: SlotStat[] = slots.map((s) => {
    const needByRole = p.roles.map((r) => p.rooms.reduce((a, m) => a + (p.need[needKey(s.key, r.id, m.id)] || 0), 0));
    return {
      need: needByRole.reduce((a, b) => a + b, 0),
      available: 0,
      assigned: 0,
      fixed: 0,
      needByRole,
      assignedByRole: p.roles.map(() => 0),
    };
  });
  const teachers: TeacherStat[] = p.teachers.map((t) => {
    const st: TeacherStat = {
      total: 0,
      fixedCount: 0,
      xCount: 0,
      maxPossible: 0,
      roleCounts: p.roles.map(() => 0),
      curLoad: 0,
      cumLoad: 0,
      dayCounts: new Array(dayCount).fill(0),
      maxConsec: 0,
      conflicts: 0,
    };
    let run = 0;
    let prevDay = -1;
    slots.forEach((s, si) => {
      const c = p.cells[cellKey(t.id, s.key)];
      if (s.dayIdx !== prevDay) {
        run = 0;
        prevDay = s.dayIdx;
      }
      if (c?.x) st.xCount++;
      else slotStats[si].available++;
      if (c?.fixed) {
        st.fixedCount++;
        slotStats[si].fixed++;
      }
      if (isAssigned(c)) {
        st.total++;
        slotStats[si].assigned++;
        st.dayCounts[s.dayIdx]++;
        run++;
        st.maxConsec = Math.max(st.maxConsec, run);
      } else run = 0;
      if (c?.role) {
        const ri = roleIdx.get(c.role);
        if (ri !== undefined) {
          hasResults = true;
          st.roleCounts[ri]++;
          st.curLoad += p.roles[ri].weight || 0;
          slotStats[si].assignedByRole[ri]++;
        }
      }
      if (c?.room && t.forbidden.includes(c.room) && isAssigned(c)) st.conflicts++;
    });
    st.maxPossible = slots.length - st.xCount;
    st.cumLoad = (t.prevLoad || 0) + st.curLoad;
    return st;
  });
  return {
    teachers,
    slots: slotStats,
    totalNeed: slotStats.reduce((a, s) => a + s.need, 0),
    totalTarget: p.teachers.reduce((a, t) => a + (t.target || 0), 0),
    hasResults,
    dayCount,
  };
}

export type IssueArea = 'info' | 'need' | 'assign' | 'result';
export interface Issue {
  level: 'error' | 'warn';
  area: IssueArea;
  msg: string;
}

export function validate(p: Project, slots: Slot[], stats: Stats): Issue[] {
  const issues: Issue[] = [];
  const add = (level: Issue['level'], area: IssueArea, msg: string) => issues.push({ level, area, msg });

  if (!p.days.length) add('error', 'info', '고사 일정이 없습니다.');
  if (!p.rooms.length) add('error', 'info', '고사실이 없습니다.');
  if (!p.teachers.length) add('error', 'info', '감독 교사가 없습니다.');
  if (!p.roles.length) add('error', 'info', '보직이 없습니다.');
  const dates = new Set<string>();
  p.days.forEach((d, i) => {
    if (!d.date) add('error', 'info', `고사 일정 ${i + 1}번째 줄의 날짜가 비어 있습니다.`);
    else if (dates.has(d.date)) add('error', 'info', `고사 날짜 ${fmtDay(d.date)}가 중복되었습니다.`);
    dates.add(d.date);
    if (d.start > d.end) add('error', 'info', `${fmtDay(d.date)}: 시작교시가 종료교시보다 큽니다.`);
  });
  const dupCheck = (names: string[], what: string) => {
    const seen = new Set<string>();
    names.forEach((n, i) => {
      const v = n.trim();
      if (!v) add('error', 'info', `${what} ${i + 1}번째 이름이 비어 있습니다.`);
      else if (seen.has(v)) add('error', 'info', `${what} 이름 "${v}"이(가) 중복되었습니다.`);
      seen.add(v);
    });
  };
  dupCheck(p.rooms.map((r) => r.name), '고사실');
  dupCheck(p.teachers.map((t) => t.name), '감독 교사');
  dupCheck(p.roles.map((r) => r.name), '보직');

  slots.forEach((s, si) => {
    const ss = stats.slots[si];
    if (ss.need > ss.available) add('error', 'need', `${slotLabel(s)}: 필요 감독 ${ss.need}명 > 감독 가능 교사 ${ss.available}명`);
    if (ss.fixed > ss.need) add('error', 'assign', `${slotLabel(s)}: 고정 배정 ${ss.fixed}명 > 필요 감독 ${ss.need}명`);
  });
  if (stats.totalNeed === 0 && slots.length && p.rooms.length) add('warn', 'need', '필요 감독 수가 아직 입력되지 않았습니다.');

  const noTarget = p.teachers.filter((t) => t.target === null).length;
  if (stats.totalNeed > 0 && stats.totalTarget !== stats.totalNeed)
    add('warn', 'assign', `배정할 시간 합계 ${stats.totalTarget} ≠ 총 필요 시간 ${stats.totalNeed}${noTarget ? ` (미입력 ${noTarget}명)` : ''}`);
  p.teachers.forEach((t, ti) => {
    const st = stats.teachers[ti];
    if (t.target === null) return;
    const target = t.target;
    if (target > st.maxPossible) add('error', 'assign', `${t.name}: 배정할 시간 ${target} > 최대 가능 시간 ${st.maxPossible}`);
    if (target < st.fixedCount) add('error', 'assign', `${t.name}: 배정할 시간 ${target} < 고정 배정 ${st.fixedCount}`);
  });

  if (stats.hasResults) {
    const roomName = new Map(p.rooms.map((r) => [r.id, r.name]));
    slots.forEach((s) => {
      const counts = new Map<string, number>();
      p.teachers.forEach((t) => {
        const c = p.cells[cellKey(t.id, s.key)];
        if (!isAssigned(c)) return;
        if (!c!.role || !c!.room) {
          add('error', 'result', `${slotLabel(s)} ${t.name}: 보직/고사실이 지정되지 않았습니다.`);
          return;
        }
        const k = `${c!.role}|${c!.room}`;
        counts.set(k, (counts.get(k) || 0) + 1);
        if (t.forbidden.includes(c!.room))
          add('error', 'result', `${slotLabel(s)} ${t.name}: 못 들어가는 고사실 ${roomName.get(c!.room) ?? ''}에 배정됨`);
      });
      p.roles.forEach((r) =>
        p.rooms.forEach((m) => {
          const need = p.need[needKey(s.key, r.id, m.id)] || 0;
          const got = counts.get(`${r.id}|${m.id}`) || 0;
          if (need !== got) add('error', 'result', `${slotLabel(s)} ${m.name} ${r.name}: 필요 ${need}명 / 배정 ${got}명`);
        }),
      );
    });
    p.teachers.forEach((t, ti) => {
      const st = stats.teachers[ti];
      if (t.target !== null && st.total !== t.target) add('warn', 'result', `${t.name}: 배정할 시간 ${t.target} / 실제 배정 ${st.total}`);
    });
  }
  return issues;
}

/** VBA 배정할시간자동채우기: 필요 시간 합계가 될 때까지 돌아가며 1시간씩 채운다 */
export function autoFillTargets(p: Project, slots: Slot[]): { project: Project; message?: string } {
  const stats = computeStats(p, slots);
  const T = p.teachers.length;
  if (!T) return { project: p, message: '감독 교사가 없습니다.' };
  const targets = stats.teachers.map((s) => s.fixedCount);
  const max = stats.teachers.map((s) => s.maxPossible);
  let sum = targets.reduce((a, b) => a + b, 0);
  if (sum > stats.totalNeed) return { project: p, message: '고정 배정 시간이 총 필요 시간보다 많습니다.' };
  const order = p.teachers.map((t, i) => i).sort((a, b) => (p.teachers[a].prevLoad || 0) - (p.teachers[b].prevLoad || 0) || a - b);
  let message: string | undefined;
  let i = 0;
  let noProgress = 0;
  while (sum < stats.totalNeed) {
    const t = order[i];
    if (targets[t] < max[t]) {
      targets[t]++;
      sum++;
      noProgress = 0;
    } else if (++noProgress >= T) {
      message = '자동 배정 시간을 모두 채울 수 없습니다. 최대 가능 시간을 확인해 주세요.';
      break;
    }
    i = (i + 1) % T;
  }
  return {
    project: { ...p, teachers: p.teachers.map((t, ti) => ({ ...t, target: targets[ti] })) },
    message,
  };
}

export function buildEngineInput(p: Project, slots: Slot[]): EngineInput {
  const roleIdx = new Map(p.roles.map((r, i) => [r.id, i]));
  const roomIdx = new Map(p.rooms.map((r, i) => [r.id, i]));
  const T = p.teachers.length;
  const S = slots.length;
  return {
    T,
    S,
    R: p.roles.length,
    M: p.rooms.length,
    slotDay: slots.map((s) => s.dayIdx),
    roleWeight: p.roles.map((r) => r.weight || 0),
    need: slots.map((s) => p.roles.map((r) => p.rooms.map((m) => Math.max(0, Math.floor(p.need[needKey(s.key, r.id, m.id)] || 0))))),
    x: p.teachers.map((t) => slots.map((s) => !!p.cells[cellKey(t.id, s.key)]?.x)),
    fixed: p.teachers.map((t) => slots.map((s) => !!p.cells[cellKey(t.id, s.key)]?.fixed)),
    fixedRole: p.teachers.map((t) =>
      slots.map((s) => {
        const c = p.cells[cellKey(t.id, s.key)];
        return c?.fixed && c.fixedRole ? (roleIdx.get(c.fixedRole) ?? -1) : -1;
      }),
    ),
    target: p.teachers.map((t) => t.target || 0),
    extra: p.teachers.map((t) => p.extra[t.id] || 0),
    forbidden: p.teachers.map((t) => p.rooms.map((m) => t.forbidden.includes(m.id))),
    prevLoad: p.teachers.map((t) => t.prevLoad || 0),
    names: {
      teacher: p.teachers.map((t) => t.name),
      slot: slots.map(slotLabel),
      role: p.roles.map((r) => r.name),
      room: p.rooms.map((r) => r.name),
    },
  };
  void roomIdx;
}

function stripResult(c: Cell | undefined): Cell | undefined {
  if (!c) return undefined;
  const { on, role, room, ...rest } = c;
  void on;
  void role;
  void room;
  return rest.x || rest.fixed ? rest : undefined;
}

export function applyGrid(p: Project, slots: Slot[], grid: (GridCell | null)[][]): Project {
  const cells: Record<string, Cell> = {};
  for (const [k, c] of Object.entries(p.cells)) {
    const s = stripResult(c);
    if (s) cells[k] = s;
  }
  p.teachers.forEach((t, ti) =>
    slots.forEach((s, si) => {
      const g = grid[ti]?.[si];
      if (!g) return;
      const k = cellKey(t.id, s.key);
      cells[k] = { ...(cells[k] || {}), on: !cells[k]?.fixed || undefined, role: p.roles[g.role].id, room: p.rooms[g.room].id };
      if (!cells[k].on) delete cells[k].on;
    }),
  );
  return { ...p, cells };
}

export function clearResults(p: Project): Project {
  const cells: Record<string, Cell> = {};
  for (const [k, c] of Object.entries(p.cells)) {
    const s = stripResult(c);
    if (s) cells[k] = s;
  }
  return { ...p, cells, lastRun: undefined };
}

/** 삭제된 일정/교사/고사실/보직을 참조하는 데이터 정리 */
export function cleanup(p: Project): Project {
  const slots = new Set(getSlots(p).map((s) => s.key));
  const roles = new Set(p.roles.map((r) => r.id));
  const rooms = new Set(p.rooms.map((r) => r.id));
  const teachers = new Set(p.teachers.map((t) => t.id));
  const need: Record<string, number> = {};
  for (const [k, v] of Object.entries(p.need)) {
    const [slot, role, room] = k.split('|');
    if (v > 0 && slots.has(slot) && roles.has(role) && rooms.has(room)) need[k] = v;
  }
  const cells: Record<string, Cell> = {};
  for (const [k, c] of Object.entries(p.cells)) {
    const [t, slot] = k.split('|');
    if (!teachers.has(t) || !slots.has(slot)) continue;
    const nc: Cell = { ...c };
    if (nc.role && !roles.has(nc.role)) delete nc.role;
    if (nc.fixedRole && !roles.has(nc.fixedRole)) delete nc.fixedRole;
    if (nc.room && !rooms.has(nc.room)) delete nc.room;
    if (nc.x || nc.fixed || nc.on) cells[k] = nc;
  }
  const extra: Record<string, number> = {};
  for (const [k, v] of Object.entries(p.extra || {})) if (teachers.has(k) && v > 0) extra[k] = v;
  return {
    ...p,
    need,
    cells,
    extra,
    teachers: p.teachers.map((t) => ({ ...t, forbidden: t.forbidden.filter((r) => rooms.has(r)) })),
  };
}

/**
 * 엑셀 양식을 다시 올렸을 때, 파일에 없는 내용(필요 감독 수·감독 불가 표시 등)은
 * 이름·날짜를 기준으로 기존 작업에서 옮겨 온다. 파일에 값이 있으면 파일 쪽을 그대로 쓴다.
 */
export function remapPreserve(prev: Project, next: Project): Project {
  const prevSlots = getSlots(prev);
  const nextSlots = getSlots(next);
  const prevSlotKey = new Map(prevSlots.map((s) => [`${s.date}|${s.period}`, s.key]));
  const prevRoleId = new Map(prev.roles.map((r) => [r.name.trim(), r.id]));
  const prevRoomId = new Map(prev.rooms.map((r) => [r.name.trim(), r.id]));
  const prevRoomName = new Map(prev.rooms.map((r) => [r.id, r.name.trim()]));
  const prevRoleName = new Map(prev.roles.map((r) => [r.id, r.name.trim()]));
  const prevTeacher = new Map(prev.teachers.map((t) => [t.name.trim(), t]));
  const nextRoomByName = new Map(next.rooms.map((r) => [r.name.trim(), r.id]));
  const nextRoleByName = new Map(next.roles.map((r) => [r.name.trim(), r.id]));

  const need = { ...next.need };
  if (!Object.keys(next.need).length) {
    nextSlots.forEach((s) => {
      const ps = prevSlotKey.get(`${s.date}|${s.period}`);
      if (!ps) return;
      next.roles.forEach((role) => {
        const pr = prevRoleId.get(role.name.trim());
        if (!pr) return;
        next.rooms.forEach((room) => {
          const pm = prevRoomId.get(room.name.trim());
          if (!pm) return;
          const v = prev.need[needKey(ps, pr, pm)];
          if (v > 0) need[needKey(s.key, role.id, room.id)] = v;
        });
      });
    });
  }

  const cells = { ...next.cells };
  if (!Object.keys(next.cells).length) {
    next.teachers.forEach((t) => {
      const pt = prevTeacher.get(t.name.trim());
      if (!pt) return;
      nextSlots.forEach((s) => {
        const ps = prevSlotKey.get(`${s.date}|${s.period}`);
        if (!ps) return;
        const c = prev.cells[cellKey(pt.id, ps)];
        if (!c) return;
        const nc: Cell = { ...c };
        const mapRole = (id?: string) => (id ? nextRoleByName.get(prevRoleName.get(id) ?? '') : undefined);
        if (nc.fixedRole) nc.fixedRole = mapRole(nc.fixedRole);
        if (nc.role) {
          const nr = mapRole(nc.role);
          const nm = nc.room ? nextRoomByName.get(prevRoomName.get(nc.room) ?? '') : undefined;
          if (nr && nm) {
            nc.role = nr;
            nc.room = nm;
          } else {
            delete nc.role;
            delete nc.room;
            if (!nc.fixed) delete nc.on;
          }
        }
        if (nc.x || nc.fixed || nc.on) cells[cellKey(t.id, s.key)] = nc;
      });
    });
  }

  const teachers = next.teachers.map((t) => {
    const pt = prevTeacher.get(t.name.trim());
    if (!pt) return t;
    return {
      ...t,
      prevLoad: t.prevLoad || pt.prevLoad || 0,
      target: t.target ?? pt.target,
      forbidden: t.forbidden.length
        ? t.forbidden
        : pt.forbidden.map((id) => nextRoomByName.get(prevRoomName.get(id) ?? '')).filter((x): x is string => !!x),
    };
  });

  return {
    ...next,
    need,
    cells,
    teachers,
    meta: { school: next.meta.school || prev.meta.school, title: next.meta.title || prev.meta.title },
  };
}

/** 다음 고사 준비: 누적 업무강도를 이월하고 일정·필요 인원·배정 결과를 비운다 */
export function nextExamProject(p: Project, stats: Stats): Project {
  return {
    ...p,
    meta: { ...p.meta, title: '' },
    days: [],
    need: {},
    cells: {},
    extra: {},
    lastRun: undefined,
    teachers: p.teachers.map((t, i) => ({ ...t, prevLoad: stats.teachers[i]?.cumLoad ?? t.prevLoad, target: null, forbidden: [] })),
  };
}

/** 교사 한 명의 개인 시간표 행 */
export function personalRows(p: Project, slots: Slot[], teacherId: string) {
  const roleById = new Map(p.roles.map((r, i) => [r.id, { r, i }]));
  const roomById = new Map(p.rooms.map((r) => [r.id, r]));
  return slots.map((s) => {
    const c = p.cells[cellKey(teacherId, s.key)];
    const assigned = isAssigned(c);
    const role = c?.role ? roleById.get(c.role) : undefined;
    return {
      slot: s,
      assigned,
      roomName: assigned && c?.room ? (roomById.get(c.room)?.name ?? '') : '',
      roleName: assigned && role ? role.r.name : assigned ? '배정(미지정)' : '',
      roleIndex: assigned && role ? role.i : -1,
    };
  });
}
