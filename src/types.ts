export interface Day {
  id: string;
  date: string; // YYYY-MM-DD
  start: number; // 시작교시
  end: number; // 종료교시
}

export interface Room {
  id: string;
  name: string;
}

export interface Role {
  id: string;
  name: string;
  weight: number; // 업무강도
}

export interface Teacher {
  id: string;
  name: string;
  prevLoad: number; // 이전 시험까지 누적 업무강도
  target: number | null; // 배정할 시간
  forbidden: string[]; // 못 들어가는 고사실(room id)
}

/** 감독배정 그리드의 한 칸 (교사 × 시험시간) */
export interface Cell {
  x?: boolean; // 감독 불가
  fixed?: boolean; // 고정 배정(반드시 이 시간에 감독)
  fixedRole?: string; // 고정 셀의 지정 보직(role id)
  on?: boolean; // 배정됨
  role?: string; // 배정 보직(role id)
  room?: string; // 배정 고사실(room id)
}

export interface Meta {
  school: string;
  title: string;
}

export interface RunInfo {
  at: string;
  trials: number;
  score: number;
  summary: string[];
}

export interface Project {
  version: 1;
  meta: Meta;
  days: Day[];
  rooms: Room[];
  roles: Role[];
  teachers: Teacher[];
  /** key: `${slotKey}|${roleId}|${roomId}` → 필요 감독 수 */
  need: Record<string, number>;
  /** key: `${teacherId}|${slotKey}` */
  cells: Record<string, Cell>;
  /** 배정 시 목표시간 초과를 허용한 교사(id → 추가 시간) */
  extra: Record<string, number>;
  lastRun?: RunInfo;
}

export interface Slot {
  key: string; // `${dayId}:${period}`
  dayId: string;
  dayIdx: number;
  date: string;
  period: number;
  posInDay: number;
  firstOfDay: boolean;
  lastOfDay: boolean;
}
