/**
 * 감독 배정 엔진 — proctor.xlsm의 VBA(감독배정풀코스)를 순수 함수로 옮기고 보강한 것.
 *
 *  1. 배정가능여부점검   : 시간 합계·교시별 가용 인원·교사별 가용 시간 점검
 *  2. 감독배정           : P값(남은 필요/남은 빈칸)이 가장 큰 교시부터, 남은 배정비율이 큰 교사에게 배정.
 *                          막히면 증가 경로(교환)를 찾아 해결하고, 그래도 안 되면 추가 배정할 교사를 묻는다.
 *  3. 날짜분산하기       : 하루에 몰린 감독을 다른 교사와 교환해 날짜별로 고르게
 *  4. 연속감독분산       : 같은 날 3연속 → 끊기, 2연속 → 비연속으로 재배치
 *  5. 보직배정           : 누적 업무강도가 낮은 교사부터 무거운 보직
 *  6. 고사실배정         : 보직별 고사실을 무작위로 섞어 배정(못 들어가는 고사실 회피)
 *  7. 배정불가고사실변경 : 같은 교시 안에서 2자/3자 교환으로 충돌 해소
 *  8. 업무분산           : 업무강도 높은 교사와 낮은 교사의 보직 교환
 *  9. 배정불가고사실변경 재실행 후 결과 지표 산출
 */

export interface EngineInput {
  T: number;
  S: number;
  R: number;
  M: number;
  slotDay: number[];
  roleWeight: number[];
  need: number[][][]; // [s][r][m]
  x: boolean[][]; // [t][s]
  fixed: boolean[][]; // [t][s]
  fixedRole: number[][]; // [t][s] -1 = 지정 없음
  target: number[];
  extra: number[];
  forbidden: boolean[][]; // [t][m]
  prevLoad: number[];
  names: { teacher: string[]; slot: string[]; role: string[]; room: string[] };
}

export interface EngineOptions {
  seed: number;
  balanceDays?: boolean;
  spreadConsecutive?: boolean;
  balanceLoad?: boolean;
}

export interface GridCell {
  role: number;
  room: number;
}

export interface Metrics {
  conflicts: number;
  consec3: number;
  consec2: number;
  dayImbalance: number;
  loadStd: number;
  loadMin: number;
  loadMax: number;
  overTarget: number;
  score: number;
}

export type EngineResult =
  | { status: 'ok'; grid: (GridCell | null)[][]; metrics: Metrics; warnings: string[] }
  | { status: 'precheck'; issues: string[] }
  | { status: 'needExtra'; slot: number; candidates: { t: number; total: number; target: number }[] }
  | { status: 'error'; message: string };

export function mulberry32(seed: number) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffle<T>(arr: T[], rnd: () => number): T[] {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(rnd() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

export function precheck(inp: EngineInput): string[] {
  const { T, S, names } = inp;
  const issues: string[] = [];
  const colNeed = inp.need.map((rr) => rr.reduce((a, row) => a + row.reduce((b, v) => b + v, 0), 0));
  const totalNeed = colNeed.reduce((a, b) => a + b, 0);
  const totalTarget = inp.target.reduce((a, b) => a + b, 0);
  if (totalNeed === 0) issues.push('필요 감독 수가 입력되지 않았습니다. [필요 감독 수] 단계에서 입력해 주세요.');
  if (totalTarget !== totalNeed)
    issues.push(`배정할 시간 합계(${totalTarget})와 총 필요 시간 합계(${totalNeed})가 일치하지 않습니다. [배정할 시간 자동 채우기]를 사용할 수 있습니다.`);
  for (let s = 0; s < S; s++) {
    let avail = 0;
    let fixedCnt = 0;
    for (let t = 0; t < T; t++) {
      if (inp.fixed[t][s]) fixedCnt++;
      else if (!inp.x[t][s]) avail++;
    }
    if (fixedCnt > colNeed[s])
      issues.push(`${names.slot[s]}: 고정 배정(${fixedCnt}명)이 필요 감독 수(${colNeed[s]}명)보다 많습니다.`);
    else if (colNeed[s] - fixedCnt > avail)
      issues.push(`${names.slot[s]}: 필요 감독 ${colNeed[s]}명인데 감독 가능한 교사가 ${avail + fixedCnt}명뿐입니다.`);
  }
  for (let t = 0; t < T; t++) {
    let open = 0;
    let fixedCnt = 0;
    for (let s = 0; s < S; s++) {
      if (inp.fixed[t][s]) fixedCnt++;
      else if (!inp.x[t][s]) open++;
    }
    const cap = inp.target[t] + inp.extra[t];
    if (fixedCnt > cap) issues.push(`${names.teacher[t]}: 고정 배정(${fixedCnt}시간)이 배정할 시간(${cap})보다 많습니다.`);
    else if (cap - fixedCnt > open)
      issues.push(`${names.teacher[t]}: 배정할 시간(${cap})이 감독 가능한 시간(${open + fixedCnt})보다 많습니다.`);
  }
  return issues;
}

export function runEngine(inp: EngineInput, opt: EngineOptions): EngineResult {
  const issues = precheck(inp);
  if (issues.length) return { status: 'precheck', issues };

  const { T, S, R, M, slotDay, roleWeight, need, forbidden, fixed, fixedRole } = inp;
  const rnd = mulberry32(opt.seed);
  const warnings: string[] = [];
  const tieKey = shuffle([...Array(T).keys()], rnd);
  const tieRank = new Array<number>(T);
  tieKey.forEach((t, i) => (tieRank[t] = i));

  // ---------- 날짜별 열 ----------
  const D = Math.max(0, ...slotDay) + 1;
  const dayCols: number[][] = Array.from({ length: D }, () => []);
  for (let s = 0; s < S; s++) dayCols[slotDay[s]].push(s);

  // ---------- 1단계: 출석(0/1/x) 그리드 ----------
  const X = -1;
  const pres: number[][] = [];
  const cnt = new Array(T).fill(0);
  const open = new Array(T).fill(0);
  const colCnt = new Array(S).fill(0);
  const colOpen = new Array(S).fill(0);
  const colNeed = need.map((rr) => rr.reduce((a, row) => a + row.reduce((b, v) => b + v, 0), 0));
  const cap = inp.target.map((v, t) => v + inp.extra[t]);

  for (let t = 0; t < T; t++) {
    pres.push(new Array(S).fill(0));
    for (let s = 0; s < S; s++) {
      if (fixed[t][s]) {
        pres[t][s] = 1;
        cnt[t]++;
        colCnt[s]++;
      } else if (inp.x[t][s]) pres[t][s] = X;
      else {
        open[t]++;
        colOpen[s]++;
      }
    }
  }
  const set = (t: number, s: number, v: 0 | 1) => {
    const old = pres[t][s];
    if (old === v) return;
    pres[t][s] = v;
    if (v === 1) {
      cnt[t]++;
      colCnt[s]++;
      open[t]--;
      colOpen[s]--;
    } else {
      cnt[t]--;
      colCnt[s]--;
      open[t]++;
      colOpen[s]++;
    }
  };

  // 증가 경로: s0에 한 명을 더 넣기 위해 교사들의 배정을 연쇄 이동
  const augment = (s0: number): boolean => {
    const parentCol = new Int32Array(S).fill(-2);
    const parentT = new Int32Array(S).fill(-1);
    parentCol[s0] = -1;
    const queue = [s0];
    while (queue.length) {
      const c = queue.shift()!;
      const order = [...Array(T).keys()].sort((a, b) => cnt[a] - cnt[b] || tieRank[a] - tieRank[b]);
      for (const t of order) {
        if (pres[t][c] !== 0) continue;
        if (cap[t] - cnt[t] > 0) {
          // 경로 적용
          let cur = c;
          set(t, cur, 1);
          while (cur !== s0) {
            const pc = parentCol[cur];
            const pt = parentT[cur];
            set(pt, cur, 0);
            set(pt, pc, 1);
            cur = pc;
          }
          return true;
        }
        for (let c2 = 0; c2 < S; c2++) {
          if (parentCol[c2] !== -2) continue;
          if (pres[t][c2] !== 1 || fixed[t][c2]) continue;
          parentCol[c2] = c;
          parentT[c2] = t;
          queue.push(c2);
        }
      }
    }
    return false;
  };

  for (let guard = 0; guard < 100000; guard++) {
    let best = -1;
    let bestP = -Infinity;
    for (let s = 0; s < S; s++) {
      const rem = colNeed[s] - colCnt[s];
      if (rem <= 0) continue;
      const p = colOpen[s] > 0 ? rem / colOpen[s] : Infinity;
      if (p > bestP) {
        bestP = p;
        best = s;
      }
    }
    if (best < 0) break;
    const s = best;
    let bt = -1;
    let bq = -Infinity;
    for (let t = 0; t < T; t++) {
      if (pres[t][s] !== 0) continue;
      const r = cap[t] - cnt[t];
      if (r <= 0) continue;
      const q = r / open[t];
      if (q > bq + 1e-12 || (Math.abs(q - bq) <= 1e-12 && tieRank[t] < tieRank[bt])) {
        bq = q;
        bt = t;
      }
    }
    if (bt >= 0) {
      set(bt, s, 1);
      continue;
    }
    if (augment(s)) continue;
    const candidates = [];
    for (let t = 0; t < T; t++) if (pres[t][s] === 0) candidates.push({ t, total: cnt[t], target: inp.target[t] });
    candidates.sort((a, b) => a.total - b.total);
    return { status: 'needExtra', slot: s, candidates };
  }

  const dayCount = (t: number, d: number) => dayCols[d].reduce((a, s) => a + (pres[t][s] === 1 ? 1 : 0), 0);
  const maxConsec = (t: number, d: number) => {
    let best = 0;
    let cur = 0;
    for (const s of dayCols[d]) {
      if (pres[t][s] === 1) {
        cur++;
        if (cur > best) best = cur;
      } else cur = 0;
    }
    return best;
  };

  // ---------- 2단계: 날짜 분산 ----------
  if (opt.balanceDays !== false && D > 1) {
    for (let guard = 0; guard < 20000; guard++) {
      let changed = false;
      outer: for (let t1 = 0; t1 < T; t1++) {
        const dc1 = Array.from({ length: D }, (_, d) => dayCount(t1, d));
        for (let hi = 0; hi < D; hi++) {
          for (let lo = 0; lo < D; lo++) {
            if (dc1[hi] - dc1[lo] <= 1) continue;
            for (let k = 1; k < T; k++) {
              const t2 = (t1 + k) % T;
              if (dayCount(t2, lo) - dayCount(t2, hi) <= 0) continue;
              const c1 = dayCols[hi].find((s) => pres[t1][s] === 1 && !fixed[t1][s] && pres[t2][s] === 0);
              const c2 = dayCols[lo].find((s) => pres[t1][s] === 0 && pres[t2][s] === 1 && !fixed[t2][s]);
              if (c1 === undefined || c2 === undefined) continue;
              set(t1, c1, 0);
              set(t2, c1, 1);
              set(t1, c2, 1);
              set(t2, c2, 0);
              changed = true;
              break outer;
            }
          }
        }
      }
      if (!changed) break;
    }
  }

  // ---------- 3단계: 연속 감독 분산 ----------
  if (opt.spreadConsecutive !== false) {
    const longestRun = (t: number, d: number) => {
      let best = 0;
      let bestStart = 0;
      let cur = 0;
      let curStart = 0;
      dayCols[d].forEach((s, p) => {
        if (pres[t][s] === 1) {
          if (cur === 0) curStart = p;
          cur++;
          if (cur > best) {
            best = cur;
            bestStart = curStart;
          }
        } else cur = 0;
      });
      return { len: best, start: bestStart };
    };

    // Phase 1: 3연속 이상
    for (let iter = 0; iter < 50; iter++) {
      let improved = false;
      for (let a = 0; a < T; a++) {
        for (let d = 0; d < D; d++) {
          if (dayCols[d].length < 3) continue;
          const run = longestRun(a, d);
          if (run.len < 3) continue;
          const sc = dayCols[d][run.start + Math.floor(run.len / 2)];
          if (fixed[a][sc]) continue;
          // 같은 날 열을 먼저, 그다음 다른 날
          const cols = [...dayCols[d], ...[...Array(S).keys()].filter((s) => slotDay[s] !== d)];
          let done = false;
          for (const b of [...Array(T).keys()].sort((p, q) => tieRank[p] - tieRank[q])) {
            if (b === a || pres[b][sc] !== 0) continue;
            for (const j of cols) {
              if (j === sc || pres[a][j] !== 0 || pres[b][j] !== 1 || fixed[b][j]) continue;
              const dj = slotDay[j];
              if (dj !== d) {
                if (dayCount(a, dj) > dayCount(a, d) - 1 - 0) continue;
                if (dayCount(b, d) > dayCount(b, dj) - 1) continue;
              }
              set(a, sc, 0);
              set(b, sc, 1);
              set(a, j, 1);
              set(b, j, 0);
              const okB = maxConsec(b, d) < 3 && maxConsec(b, dj) < 3;
              const okA = maxConsec(a, dj) < 3 && maxConsec(a, d) < run.len;
              if (okA && okB) {
                done = true;
                break;
              }
              set(a, sc, 1);
              set(b, sc, 0);
              set(a, j, 0);
              set(b, j, 1);
            }
            if (done) break;
          }
          if (done) improved = true;
        }
      }
      if (!improved) break;
    }

    // Phase 2: 정확히 2연속 → 같은 날 비연속으로
    for (let iter = 0; iter < 30; iter++) {
      let improved = false;
      for (let a = 0; a < T; a++) {
        for (let d = 0; d < D; d++) {
          if (dayCols[d].length < 3) continue;
          const run = longestRun(a, d);
          if (run.len !== 2) continue;
          let pos = run.start;
          let sc = dayCols[d][pos];
          if (fixed[a][sc]) {
            pos = run.start + 1;
            sc = dayCols[d][pos];
            if (fixed[a][sc]) continue;
          }
          let done = false;
          for (const q of dayCols[d]) {
            if (q === sc || pres[a][q] !== 0) continue;
            set(a, sc, 0);
            set(a, q, 1);
            const aRun = maxConsec(a, d);
            set(a, sc, 1);
            set(a, q, 0);
            if (aRun >= 2) continue;
            for (let b = 0; b < T; b++) {
              if (b === a || pres[b][sc] !== 0 || pres[b][q] !== 1 || fixed[b][q]) continue;
              set(b, sc, 1);
              set(b, q, 0);
              const bRun = maxConsec(b, d);
              set(b, sc, 0);
              set(b, q, 1);
              if (bRun >= 2) continue;
              set(a, sc, 0);
              set(b, sc, 1);
              set(a, q, 1);
              set(b, q, 0);
              done = true;
              break;
            }
            if (done) break;
          }
          if (done) improved = true;
        }
      }
      if (!improved) break;
    }
  }

  // ---------- 4단계: 보직 배정 ----------
  const role: number[][] = Array.from({ length: T }, () => new Array(S).fill(-1));
  const room: number[][] = Array.from({ length: T }, () => new Array(S).fill(-1));
  const cur = new Array(T).fill(0);
  const cum = () => inp.prevLoad.map((p, t) => p + cur[t]);

  for (let s = 0; s < S; s++) {
    const remain = need[s].map((row) => row.reduce((a, v) => a + v, 0));
    for (let t = 0; t < T; t++) {
      if (!fixed[t][s] || fixedRole[t][s] < 0) continue;
      const fr = fixedRole[t][s];
      if (remain[fr] > 0) {
        role[t][s] = fr;
        remain[fr]--;
        cur[t] += roleWeight[fr];
      } else {
        warnings.push(`${inp.names.teacher[t]} ${inp.names.slot[s]}: 지정한 보직(${inp.names.role[fr]})의 필요 인원을 초과해 자동 배정했습니다.`);
      }
    }
    for (let r = 0; r < R; r++) {
      if (remain[r] <= 0) continue;
      const loads = cum();
      const cands = [];
      for (let t = 0; t < T; t++) if (pres[t][s] === 1 && role[t][s] < 0) cands.push(t);
      cands.sort((a, b) => loads[a] - loads[b] || tieRank[a] - tieRank[b]);
      if (cands.length < remain[r]) {
        return { status: 'error', message: `${inp.names.slot[s]} ${inp.names.role[r]}: 배정 가능한 교사가 ${remain[r] - cands.length}명 부족합니다.` };
      }
      for (let i = 0; i < remain[r]; i++) {
        role[cands[i]][s] = r;
        cur[cands[i]] += roleWeight[r];
      }
    }
  }

  // ---------- 5단계: 고사실 배정 ----------
  for (let s = 0; s < S; s++) {
    for (let r = 0; r < R; r++) {
      const rooms: number[] = [];
      for (let m = 0; m < M; m++) for (let k = 0; k < need[s][r][m]; k++) rooms.push(m);
      shuffle(rooms, rnd);
      const ts = [];
      for (let t = 0; t < T; t++) if (role[t][s] === r) ts.push(t);
      const forbCount = (t: number) => rooms.filter((m) => forbidden[t][m]).length;
      ts.sort((a, b) => forbCount(b) - forbCount(a));
      for (const t of ts) {
        let idx = rooms.findIndex((m) => !forbidden[t][m]);
        if (idx < 0) idx = 0;
        room[t][s] = rooms[idx];
        rooms.splice(idx, 1);
      }
    }
  }

  const canHold = (t: number, s: number, r: number, m: number) =>
    !forbidden[t][m] && (fixedRole[t][s] < 0 || !fixed[t][s] || fixedRole[t][s] === r);
  const swap2 = (a: number, b: number, s: number) => {
    [role[a][s], role[b][s]] = [role[b][s], role[a][s]];
    [room[a][s], room[b][s]] = [room[b][s], room[a][s]];
  };

  // ---------- 6단계: 못 들어가는 고사실 충돌 해소 ----------
  const fixConflicts = () => {
    for (let pass = 0; pass < 10; pass++) {
      let changed = false;
      for (let s = 0; s < S; s++) {
        for (let i = 0; i < T; i++) {
          if (role[i][s] < 0 || canHold(i, s, role[i][s], room[i][s])) continue;
          const ir = role[i][s];
          let swapped = false;
          for (let diff = 0; diff < Math.max(1, R) && !swapped; diff++) {
            // 2자간
            for (let k = 1; k < T; k++) {
              const u = (i + k) % T;
              const ur = role[u][s];
              if (ur < 0 || Math.abs(ur - ir) !== diff) continue;
              if (canHold(i, s, ur, room[u][s]) && canHold(u, s, ir, room[i][s])) {
                swap2(i, u, s);
                swapped = changed = true;
                break;
              }
            }
            if (swapped) break;
            // 3자간 순환
            for (let jj = 1; jj < T && !swapped; jj++) {
              const j = (i + jj) % T;
              const jr = role[j][s];
              if (jr < 0) continue;
              for (let kk = jj + 1; kk < T; kk++) {
                const k = (i + kk) % T;
                const kr = role[k][s];
                if (kr < 0) continue;
                if (Math.max(ir, jr, kr) - Math.min(ir, jr, kr) > diff) continue;
                const ci = { r: ir, m: room[i][s] };
                const cj = { r: jr, m: room[j][s] };
                const ck = { r: kr, m: room[k][s] };
                // 회전1: i←j, j←k, k←i
                if (canHold(i, s, cj.r, cj.m) && canHold(j, s, ck.r, ck.m) && canHold(k, s, ci.r, ci.m)) {
                  role[i][s] = cj.r; room[i][s] = cj.m;
                  role[j][s] = ck.r; room[j][s] = ck.m;
                  role[k][s] = ci.r; room[k][s] = ci.m;
                  swapped = changed = true;
                  break;
                }
                // 회전2: i←k, k←j, j←i
                if (canHold(i, s, ck.r, ck.m) && canHold(k, s, cj.r, cj.m) && canHold(j, s, ci.r, ci.m)) {
                  role[i][s] = ck.r; room[i][s] = ck.m;
                  role[k][s] = cj.r; room[k][s] = cj.m;
                  role[j][s] = ci.r; room[j][s] = ci.m;
                  swapped = changed = true;
                  break;
                }
              }
            }
          }
        }
      }
      if (!changed) break;
    }
  };
  fixConflicts();

  const recomputeLoad = () => {
    cur.fill(0);
    for (let t = 0; t < T; t++) for (let s = 0; s < S; s++) if (role[t][s] >= 0) cur[t] += roleWeight[role[t][s]];
  };
  recomputeLoad();

  // ---------- 7단계: 업무강도 분산 ----------
  if (opt.balanceLoad !== false && R > 1) {
    for (let guard = 0; guard < 20000; guard++) {
      const load = cum();
      const order = [...Array(T).keys()].sort((a, b) => load[b] - load[a] || a - b);
      let found = false;
      search: for (let i = 0; i < T; i++) {
        for (let j = T - 1; j > i; j--) {
          const a = order[i];
          const b = order[j];
          const diff = load[a] - load[b];
          if (diff <= 1e-9) continue;
          let best = diff - 1e-9;
          let bestS = -1;
          for (let s = 0; s < S; s++) {
            const ra = role[a][s];
            const rb = role[b][s];
            if (ra < 0 || rb < 0) continue;
            const wa = roleWeight[ra];
            const wb = roleWeight[rb];
            if (wa <= wb) continue;
            if (!canHold(a, s, rb, room[b][s]) || !canHold(b, s, ra, room[a][s])) continue;
            const fut = Math.abs(load[a] - wa + wb - (load[b] - wb + wa));
            if (fut < best) {
              best = fut;
              bestS = s;
            }
          }
          if (bestS >= 0) {
            swap2(a, b, bestS);
            recomputeLoad();
            found = true;
            break search;
          }
        }
      }
      if (!found) break;
    }
  }

  fixConflicts();
  recomputeLoad();

  // ---------- 결과 ----------
  const grid: (GridCell | null)[][] = [];
  let conflicts = 0;
  for (let t = 0; t < T; t++) {
    grid.push([]);
    for (let s = 0; s < S; s++) {
      if (pres[t][s] === 1 && role[t][s] >= 0) {
        grid[t].push({ role: role[t][s], room: room[t][s] });
        if (forbidden[t][room[t][s]]) conflicts++;
      } else grid[t].push(null);
    }
  }
  const metrics = computeMetrics(inp, grid, conflicts, dayCols);
  return { status: 'ok', grid, metrics, warnings };
}

function computeMetrics(inp: EngineInput, grid: (GridCell | null)[][], conflicts: number, dayCols: number[][]): Metrics {
  const { T } = inp;
  let consec3 = 0;
  let consec2 = 0;
  let dayImbalance = 0;
  let overTarget = 0;
  const loads: number[] = [];
  for (let t = 0; t < T; t++) {
    let total = 0;
    let load = inp.prevLoad[t];
    const counts: number[] = [];
    dayCols.forEach((cols) => {
      let run = 0;
      let best = 0;
      let c = 0;
      let possible = false;
      for (const s of cols) {
        if (!inp.x[t][s]) possible = true;
        const g = grid[t][s];
        if (g) {
          c++;
          run++;
          best = Math.max(best, run);
          load += inp.roleWeight[g.role];
        } else run = 0;
      }
      total += c;
      if (possible) counts.push(c);
      if (best >= 3) consec3++;
      else if (best === 2) consec2++;
    });
    if (counts.length > 1) dayImbalance += Math.max(0, Math.max(...counts) - Math.min(...counts) - 1);
    if (total > inp.target[t]) overTarget++;
    loads.push(load);
  }
  const mean = loads.reduce((a, b) => a + b, 0) / Math.max(1, T);
  const loadStd = Math.sqrt(loads.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, T));
  const score = conflicts * 1e6 + consec3 * 1000 + dayImbalance * 200 + consec2 * 10 + loadStd;
  return {
    conflicts,
    consec3,
    consec2,
    dayImbalance,
    loadStd,
    loadMin: T ? Math.min(...loads) : 0,
    loadMax: T ? Math.max(...loads) : 0,
    overTarget,
    score,
  };
}

/** 여러 번 시도해 점수가 가장 좋은 결과를 고른다 */
export function runBest(inp: EngineInput, trials: number, baseSeed = Date.now()): EngineResult & { trialsRun?: number } {
  let best: EngineResult | null = null;
  let run = 0;
  for (let i = 0; i < Math.max(1, trials); i++) {
    const res = runEngine(inp, { seed: (baseSeed + i * 7919) >>> 0 });
    run++;
    if (res.status !== 'ok') return res;
    if (!best || (best.status === 'ok' && res.metrics.score < best.metrics.score)) best = res;
    if (res.metrics.score < 1) break;
  }
  return { ...(best as EngineResult), trialsRun: run };
}
