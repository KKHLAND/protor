import { useEffect, useMemo, useRef, useState, type KeyboardEvent as RKE, type MouseEvent as RME } from 'react';
import { IssueList, Modal, NumberInput, Popover, useApp } from '../ui';
import { applyGrid, autoFillTargets, buildEngineInput, cellKey, clearResults, dayAllowed, fmtDay, fmtMD, isAssigned, roleColor, slotLabel } from '../model';
import { runBest, type EngineResult, type Metrics } from '../engine';
import type { Cell, Project, Teacher } from '../types';

const sk = (t: number, s: number) => `${t},${s}`;

function compact(c: Cell): Cell | undefined {
  const n: Cell = {};
  if (c.x) return { x: true };
  if (c.fixed) n.fixed = true;
  if (c.fixed && c.fixedRole) n.fixedRole = c.fixedRole;
  if (c.on && !c.fixed) n.on = true;
  if ((c.on || c.fixed) && c.role) n.role = c.role;
  if ((c.on || c.fixed) && c.room) n.room = c.room;
  return n.fixed || n.on ? n : undefined;
}

export default function AssignView() {
  const { store, slots, stats, issues, notify, confirm, setTab, openFile, downloadTemplate } = useApp();
  const p = store.project;
  const T = p.teachers.length;
  const S = slots.length;
  const roleIdx = useMemo(() => new Map(p.roles.map((r, i) => [r.id, i])), [p.roles]);
  const roomName = useMemo(() => new Map(p.rooms.map((r) => [r.id, r.name])), [p.rooms]);
  const dayById = useMemo(() => new Map(p.days.map((d) => [d.id, d])), [p.days]);
  const condLabel = (t: Teacher) => {
    const parts: string[] = [];
    if (t.availableDays?.length) {
      const names = t.availableDays.map((id) => fmtMD(dayById.get(id)?.date ?? ''));
      parts.push(names.length <= 2 ? `${names.join(',')}만` : `${names.length}일만`);
    }
    if (t.forbidden.length) parts.push(`✕${t.forbidden.map((id) => roomName.get(id)).join(',')}`);
    return parts.length ? parts.join(' · ') : '—';
  };

  const [sel, setSel] = useState<Set<string>>(new Set());
  const anchor = useRef<[number, number] | null>(null);
  const dragging = useRef(false);
  const dragBase = useRef<Set<string>>(new Set());
  const [editor, setEditor] = useState<{ t: number; s: number; x: number; y: number } | null>(null);
  const [forb, setForb] = useState<{ t: number; x: number; y: number } | null>(null);
  const [trials, setTrials] = useState(20);
  const [running, setRunning] = useState(false);
  const [pre, setPre] = useState<string[] | null>(null);
  const [needExtra, setNeedExtra] = useState<{ slot: number; candidates: { t: number; total: number; target: number }[]; extra: Record<string, number> } | null>(null);
  const [pick, setPick] = useState<number | null>(null);
  const [result, setResult] = useState<{ metrics: Metrics; warnings: string[]; ms: number; trials: number } | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const up = () => (dragging.current = false);
    window.addEventListener('mouseup', up);
    return () => window.removeEventListener('mouseup', up);
  }, []);

  const cellAt = (t: number, s: number): Cell | undefined => p.cells[cellKey(p.teachers[t].id, slots[s].key)];
  const rect = (a: [number, number], b: [number, number]) => {
    const out = new Set<string>();
    for (let t = Math.min(a[0], b[0]); t <= Math.max(a[0], b[0]); t++) for (let s = Math.min(a[1], b[1]); s <= Math.max(a[1], b[1]); s++) out.add(sk(t, s));
    return out;
  };
  const selected = (): [number, number][] =>
    [...sel].map((k) => k.split(',').map(Number) as [number, number]).filter(([t, s]) => t < T && s < S);

  const mutate = (keys: [number, number][], fn: (c: Cell | undefined, t: number, s: number) => Cell | undefined) => {
    if (!keys.length) return;
    store.update((pp) => {
      const cells = { ...pp.cells };
      for (const [t, s] of keys) {
        const k = cellKey(pp.teachers[t].id, slots[s].key);
        const n = fn(cells[k], t, s);
        const c = n ? compact(n) : undefined;
        if (c) cells[k] = c;
        else delete cells[k];
      }
      return { ...pp, cells };
    });
  };

  const setX = () => {
    const keys = selected();
    const all = keys.length > 0 && keys.every(([t, s]) => cellAt(t, s)?.x);
    mutate(keys, () => (all ? undefined : { x: true }));
  };
  const setFixed = () => {
    const keys = selected();
    const all = keys.length > 0 && keys.every(([t, s]) => cellAt(t, s)?.fixed);
    mutate(keys, (c) => (all ? { ...c, fixed: false, on: !!c?.role } : { ...c, x: false, fixed: true }));
  };
  const clearSel = () => mutate(selected(), () => undefined);

  const swap = () => {
    const keys = selected();
    let pairs: [[number, number], [number, number]][] = [];
    if (keys.length === 2) pairs = [[keys[0], keys[1]]];
    else if (keys.length === 4) {
      const ts = [...new Set(keys.map((k) => k[0]))];
      const ss = [...new Set(keys.map((k) => k[1]))];
      if (ts.length === 2 && ss.length === 2)
        pairs = [
          [[ts[0], ss[0]], [ts[1], ss[0]]],
          [[ts[0], ss[1]], [ts[1], ss[1]]],
        ];
    }
    if (!pairs.length) {
      notify('교환할 두 칸(또는 두 교사 × 두 시간의 4칸)을 선택하세요.', 'warn');
      return;
    }
    if (keys.some(([t, s]) => cellAt(t, s)?.x)) {
      notify('감독 불가(X) 칸은 교환할 수 없습니다.', 'warn');
      return;
    }
    store.update((pp) => {
      const cells = { ...pp.cells };
      for (const [[ta, sa], [tb, sb]] of pairs) {
        const ka = cellKey(pp.teachers[ta].id, slots[sa].key);
        const kb = cellKey(pp.teachers[tb].id, slots[sb].key);
        const a = cells[ka];
        const b = cells[kb];
        const content = (c?: Cell) => ({ on: isAssigned(c), role: c?.role, room: c?.room });
        const na = compact({ fixed: a?.fixed, fixedRole: a?.fixedRole, ...content(b) });
        const nb = compact({ fixed: b?.fixed, fixedRole: b?.fixedRole, ...content(a) });
        if (na) cells[ka] = na;
        else delete cells[ka];
        if (nb) cells[kb] = nb;
        else delete cells[kb];
      }
      return { ...pp, cells };
    });
    notify('교환했습니다. 아래 점검 결과를 확인하세요.', 'ok');
  };

  const onCellDown = (e: RME, t: number, s: number) => {
    if (e.button !== 0) return;
    e.preventDefault();
    gridRef.current?.focus({ preventScroll: true });
    const multi = e.ctrlKey || e.metaKey;
    if (e.shiftKey && anchor.current) {
      const r = rect(anchor.current, [t, s]);
      dragBase.current = multi ? new Set(sel) : new Set();
      setSel(new Set([...dragBase.current, ...r]));
      dragging.current = true;
      return;
    }
    if (multi) {
      const n = new Set(sel);
      if (n.has(sk(t, s))) n.delete(sk(t, s));
      else n.add(sk(t, s));
      setSel(n);
      dragBase.current = n;
    } else {
      dragBase.current = new Set();
      setSel(new Set([sk(t, s)]));
    }
    anchor.current = [t, s];
    dragging.current = true;
  };
  const onCellEnter = (t: number, s: number) => {
    if (!dragging.current || !anchor.current) return;
    setSel(new Set([...dragBase.current, ...rect(anchor.current, [t, s])]));
  };
  const selectRange = (t0: number, t1: number, s0: number, s1: number, e: RME) => {
    const r = rect([t0, s0], [t1, s1]);
    setSel(e.ctrlKey || e.metaKey ? new Set([...sel, ...r]) : r);
    anchor.current = [t0, s0];
    gridRef.current?.focus({ preventScroll: true });
  };

  const onKeyDown = (e: RKE) => {
    if (editor || forb) return;
    const code = e.code;
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && code === 'KeyC') {
      e.preventDefault();
      swap();
      return;
    }
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (code === 'KeyX') {
      e.preventDefault();
      setX();
    } else if (code === 'Digit1' || code === 'Numpad1' || code === 'KeyF') {
      e.preventDefault();
      setFixed();
    } else if (code === 'Delete' || code === 'Backspace') {
      e.preventDefault();
      clearSel();
    } else if (code === 'Escape') setSel(new Set());
    else if (code === 'Enter' && anchor.current) {
      e.preventDefault();
      const [t, s] = anchor.current;
      const el = gridRef.current?.querySelector<HTMLElement>(`td[data-k="${sk(t, s)}"]`);
      const r = el?.getBoundingClientRect();
      setEditor({ t, s, x: r ? r.left : 200, y: r ? r.bottom : 200 });
    } else if (code.startsWith('Arrow') && T && S) {
      e.preventDefault();
      const [t, s] = anchor.current ?? [0, 0];
      const nt = Math.max(0, Math.min(T - 1, t + (code === 'ArrowDown' ? 1 : code === 'ArrowUp' ? -1 : 0)));
      const ns = Math.max(0, Math.min(S - 1, s + (code === 'ArrowRight' ? 1 : code === 'ArrowLeft' ? -1 : 0)));
      if (e.shiftKey) setSel(rect([t, s], [nt, ns]));
      else {
        setSel(new Set([sk(nt, ns)]));
        anchor.current = [nt, ns];
      }
      gridRef.current?.querySelector(`td[data-k="${sk(nt, ns)}"]`)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  };

  const run = async (extra: Record<string, number>) => {
    const blocking = issues.filter((i) => i.level === 'error' && (i.area === 'info' || i.area === 'need'));
    if (blocking.length) {
      setPre(blocking.map((i) => i.msg));
      return;
    }
    const base = clearResults({ ...p, extra });
    const inp = buildEngineInput(base, slots);
    setRunning(true);
    await new Promise((r) => setTimeout(r, 30));
    const t0 = performance.now();
    let res: EngineResult & { trialsRun?: number };
    try {
      res = runBest(inp, trials);
    } catch (err) {
      setRunning(false);
      notify(`배정 중 오류가 발생했습니다: ${(err as Error).message}`, 'error');
      return;
    }
    setRunning(false);
    if (res.status === 'precheck') {
      setPre(res.issues);
      return;
    }
    if (res.status === 'needExtra') {
      setPick(null);
      setNeedExtra({ slot: res.slot, candidates: res.candidates, extra });
      return;
    }
    if (res.status === 'error') {
      notify(res.message, 'error');
      return;
    }
    const ok = res;
    const m = ok.metrics;
    const summary = [
      `못 들어가는 고사실 충돌 ${m.conflicts}건`,
      `같은 날 3연속 이상 ${m.consec3}건 · 2연속 ${m.consec2}건`,
      `누적 업무강도 ${m.loadMin}~${m.loadMax} (표준편차 ${m.loadStd.toFixed(1)})`,
    ];
    store.update((pp: Project) => ({
      ...applyGrid(clearResults({ ...pp, extra }), slots, ok.grid),
      lastRun: { at: new Date().toISOString(), trials: ok.trialsRun ?? trials, score: m.score, summary },
    }));
    setSel(new Set());
    setResult({ metrics: m, warnings: ok.warnings, ms: performance.now() - t0, trials: ok.trialsRun ?? trials });
  };

  const startRun = async () => {
    if (stats.hasResults && !(await confirm({ title: '감독 배정 시작', message: '기존 배정 결과를 지우고 새로 배정합니다. (감독 불가·고정 표시는 유지됩니다)', ok: '새로 배정' }))) return;
    run(p.extra);
  };

  const autoFill = async () => {
    const lockedCount = p.teachers.filter((t) => t.lockTarget).length;
    if (
      p.teachers.some((t) => t.target !== null) &&
      !(await confirm({
        title: '배정할 시간 자동 채우기',
        message: (
          <>
            <p>배정할 시간을 새로 계산합니다. 이전 누적 업무강도가 낮은 교사부터 1시간씩 돌아가며 채웁니다.</p>
            {lockedCount > 0 && <p className="muted">🔒 시수를 고정한 {lockedCount}명은 그대로 두고, 남은 시간만 나머지 교사에게 나눕니다.</p>}
          </>
        ),
        ok: '자동 채우기',
      }))
    )
      return;
    const res = autoFillTargets(p, slots);
    store.update(() => res.project);
    if (res.message) notify(res.message, 'warn');
    else notify('배정할 시간을 채웠습니다.', 'ok');
  };

  if (!T || !S) {
    return (
      <div className="page">
        <div className="page-head">
          <div>
            <h2>감독 배정</h2>
          </div>
        </div>
        <div className="empty-state">
          <p>먼저 기본 정보에서 고사 일정과 감독 교사를 입력하세요.</p>
          <button className="btn primary" onClick={() => setTab('info')}>
            기본 정보로 이동
          </button>
        </div>
      </div>
    );
  }

  const dayGroups: { dayIdx: number; date: string; start: number; count: number }[] = [];
  slots.forEach((s, i) => {
    const g = dayGroups[dayGroups.length - 1];
    if (g && g.dayIdx === s.dayIdx) g.count++;
    else dayGroups.push({ dayIdx: s.dayIdx, date: s.date, start: i, count: 1 });
  });

  const assignIssues = issues.filter((i) => i.area === 'assign' || i.area === 'result' || i.area === 'need');
  const errCount = assignIssues.filter((i) => i.level === 'error').length;
  const sumOk = stats.totalTarget === stats.totalNeed && stats.totalNeed > 0;
  const extras = Object.entries(p.extra).filter(([, v]) => v > 0);
  const selCount = sel.size;

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h2>감독 배정</h2>
          <p className="sub">
            감독할 수 없는 시간은 <kbd>X</kbd>, 반드시 감독할 시간은 <kbd>1</kbd>(고정)로 표시하고 교사별 “못 들어가는 고사실”을 지정한 뒤, 아래 순서대로
            누르세요.
          </p>
        </div>
      </div>

      <section className="card toolbar-card">
        <div className="toolbar">
          <button className="btn" onClick={autoFill}>
            <span className="num-badge">1</span> 배정할 시간 자동 채우기
          </button>
          <button className="btn primary" onClick={startRun} disabled={running}>
            <span className="num-badge light">2</span> {running ? '배정하는 중…' : '감독 배정 시작'}
          </button>
          <label className="inline">
            시도
            <select className="inp" value={trials} onChange={(e) => setTrials(Number(e.target.value))} title="여러 번 배정해 가장 좋은 결과를 고릅니다">
              {[1, 10, 20, 50, 100].map((n) => (
                <option key={n} value={n}>
                  {n}회
                </option>
              ))}
            </select>
          </label>
          <button className="btn accent" onClick={() => setTab('chart')} disabled={!stats.hasResults}>
            <span className="num-badge light">3</span> 감독표 보기
          </button>
          <span className="sep" />
          <button className="btn sm" onClick={downloadTemplate} title="교사 명단과 교시가 채워진 엑셀 양식을 내려받습니다">
            엑셀 양식 내려받기
          </button>
          <button className="btn sm" onClick={openFile}>
            작성한 엑셀 올리기
          </button>
          <div className="spacer" />
          <button
            className="btn sm"
            disabled={!stats.hasResults}
            onClick={async () => {
              if (await confirm({ title: '배정 결과 초기화', message: '자동 배정 결과를 지웁니다. 감독 불가(X)와 고정 표시는 유지됩니다.', ok: '결과 지우기' })) store.update(clearResults);
            }}
          >
            배정 결과 초기화
          </button>
          <button
            className="btn sm danger-outline"
            onClick={async () => {
              if (await confirm({ title: '표 전체 초기화', message: '감독 불가·고정 표시와 배정 결과를 모두 지웁니다.', ok: '전체 지우기', danger: true }))
                store.update((pp) => ({ ...pp, cells: {}, extra: {}, lastRun: undefined }));
            }}
          >
            표 전체 초기화
          </button>
        </div>
        <div className="toolbar second">
          <span className="tool-label">선택 {selCount ? `${selCount}칸` : ''}</span>
          <button className="btn sm" onClick={setX} disabled={!selCount} title="X">
            ✕ 감독 불가
          </button>
          <button className="btn sm" onClick={setFixed} disabled={!selCount} title="1 또는 F">
            📌 고정
          </button>
          <button className="btn sm" onClick={clearSel} disabled={!selCount} title="Delete">
            지우기
          </button>
          <button className="btn sm" onClick={swap} disabled={selCount !== 2 && selCount !== 4} title="Ctrl+Shift+C">
            ⇄ 교환
          </button>
          <span className="hint-inline">드래그·Shift/Ctrl+클릭으로 선택 · 머리글 클릭으로 줄 선택 · 더블클릭으로 칸 편집 · 🔒 시수 고정 · [근무 조건]에서 감독 가능 날짜 지정</span>
          <div className="spacer" />
          <span className={`chip ${sumOk ? 'ok' : 'err'}`}>
            배정할 시간 합 {stats.totalTarget} / 총 필요 시간 {stats.totalNeed}
          </span>
          {extras.map(([id, v]) => (
            <span key={id} className="chip warn">
              목표 초과 허용: {p.teachers.find((t) => t.id === id)?.name} +{v}
              <button
                className="chip-x"
                aria-label="해제"
                onClick={() =>
                  store.update((pp) => {
                    const extra = { ...pp.extra };
                    delete extra[id];
                    return { ...pp, extra };
                  })
                }
              >
                ✕
              </button>
            </span>
          ))}
        </div>
      </section>

      <div className="legend">
        {p.roles.map((r, i) => (
          <span key={r.id}>
            <i className="swatch" style={{ background: roleColor(i) }} />
            {r.name}
          </span>
        ))}
        <span>
          <i className="swatch x" />
          감독 불가
        </span>
        <span>
          <i className="swatch blocked" />
          감독 가능 날짜 아님
        </span>
        <span>
          <i className="swatch fixed" />
          고정
        </span>
        <span>
          <i className="swatch conflict" />
          못 들어가는 고사실
        </span>
        {p.lastRun && <span className="muted last-run">마지막 배정 {new Date(p.lastRun.at).toLocaleString('ko-KR')} · {p.lastRun.summary.join(' · ')}</span>}
      </div>

      <div className="agrid-wrap" ref={gridRef} tabIndex={0} onKeyDown={onKeyDown}>
        <table className="agrid">
          <thead>
            <tr>
              <th className="sticky c-no" rowSpan={2}>
                순번
              </th>
              <th className="sticky c-name" rowSpan={2}>
                이름
              </th>
              <th className="sticky c-target" rowSpan={2}>
                배정할
                <br />
                시간
              </th>
              <th className="sticky c-forb" rowSpan={2}>
                근무 조건
                <br />
                (날짜·고사실)
              </th>
              {dayGroups.map((g) => (
                <th key={g.dayIdx} colSpan={g.count} className="day-head day-start clickable" onMouseDown={(e) => selectRange(0, T - 1, g.start, g.start + g.count - 1, e)}>
                  {fmtDay(g.date)}
                </th>
              ))}
              <th rowSpan={2} className="st st-first">
                최대
                <br />
                가능
              </th>
              <th rowSpan={2} className="st">
                총감독
                <br />
                시간
              </th>
              <th rowSpan={2} className="st">
                누적
                <br />
                업무강도
              </th>
              {p.roles.map((r, i) => (
                <th key={r.id} rowSpan={2} className="st" style={{ boxShadow: `inset 0 -3px 0 ${roleColor(i)}` }}>
                  {r.name}
                </th>
              ))}
              <th rowSpan={2} className="st">
                현재
                <br />
                업무강도
              </th>
            </tr>
            <tr>
              {slots.map((s, si) => (
                <th key={s.key} className={`per-head clickable ${s.firstOfDay ? 'day-start' : ''}`} onMouseDown={(e) => selectRange(0, T - 1, si, si, e)}>
                  {s.period}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {p.teachers.map((t, ti) => {
              const st = stats.teachers[ti];
              const target = t.target;
              return (
                <tr key={t.id}>
                  <td className="sticky c-no muted">{ti + 1}</td>
                  <td className="sticky c-name clickable" onMouseDown={(e) => selectRange(ti, ti, 0, S - 1, e)} title={t.name}>
                    {t.name}
                  </td>
                  <td className={`sticky c-target ${t.lockTarget ? 'locked' : ''} ${target !== null && (target > st.maxPossible || target < st.fixedCount) ? 'bad' : ''}`}>
                    <div className="target-cell">
                      <button
                        className={`lock ${t.lockTarget ? 'on' : ''}`}
                        title={t.lockTarget ? '시수 고정됨 — [배정할 시간 자동 채우기]가 이 값을 바꾸지 않습니다' : '시수 고정하기 (강사처럼 시수가 정해진 교사)'}
                        onClick={() =>
                          store.update((pp) => ({ ...pp, teachers: pp.teachers.map((x) => (x.id === t.id ? { ...x, lockTarget: x.lockTarget ? undefined : true } : x)) }))
                        }
                      >
                        {t.lockTarget ? '🔒' : '🔓'}
                      </button>
                      <NumberInput
                        integer
                        min={0}
                        allowEmpty
                        value={target}
                        onChange={(v) => store.update((pp) => ({ ...pp, teachers: pp.teachers.map((x) => (x.id === t.id ? { ...x, target: v } : x)) }), { tag: `target:${t.id}` })}
                      />
                    </div>
                  </td>
                  <td className="sticky c-forb">
                    <button
                      className={`forb-btn ${t.forbidden.length || t.availableDays?.length ? 'has' : ''}`}
                      title="감독 가능 날짜와 못 들어가는 고사실을 설정합니다"
                      onClick={(e) => setForb({ t: ti, x: e.clientX, y: e.clientY })}
                    >
                      {condLabel(t)}
                    </button>
                  </td>
                  {slots.map((s, si) => {
                    const c = p.cells[cellKey(t.id, s.key)];
                    const ri = c?.role ? (roleIdx.get(c.role) ?? -1) : -1;
                    const assigned = isAssigned(c);
                    const conflict = assigned && !!c?.room && t.forbidden.includes(c.room);
                    const blocked = !dayAllowed(t, s) && !c?.fixed;
                    const k = sk(ti, si);
                    const cls = [
                      'cell',
                      s.firstOfDay ? 'day-start' : '',
                      blocked ? 'blocked' : '',
                      c?.x ? 'x' : '',
                      c?.fixed ? 'fixed' : '',
                      assigned ? 'on' : '',
                      ri >= 0 ? 'role' : '',
                      sel.has(k) ? 'sel' : '',
                      conflict ? 'conflict' : '',
                    ].join(' ');
                    let text = '';
                    if (c?.x) text = 'X';
                    else if (blocked && !assigned) text = '·';
                    else if (assigned && c?.room) text = roomName.get(c.room) ?? '?';
                    else if (c?.fixed) text = c.fixedRole ? `고정·${p.roles[roleIdx.get(c.fixedRole) ?? 0]?.name.slice(0, 1) ?? ''}` : '고정';
                    else if (assigned) text = '1';
                    const title = `${t.name} · ${slotLabel(s)}${ri >= 0 ? ` · ${roomName.get(c!.room!) ?? ''} ${p.roles[ri].name}` : ''}${c?.fixed ? ' · 고정' : ''}${blocked ? ' · 감독 가능 날짜가 아님' : ''}${conflict ? ' · 못 들어가는 고사실!' : ''}`;
                    return (
                      <td
                        key={s.key}
                        data-k={k}
                        className={cls}
                        style={ri >= 0 && !c?.fixed ? { background: roleColor(ri) } : ri >= 0 ? { background: `linear-gradient(90deg, #d4d4d4 0 4px, ${roleColor(ri)} 4px)` } : undefined}
                        title={title}
                        onMouseDown={(e) => onCellDown(e, ti, si)}
                        onMouseEnter={() => onCellEnter(ti, si)}
                        onDoubleClick={(e) => setEditor({ t: ti, s: si, x: e.clientX, y: e.clientY })}
                      >
                        {text}
                      </td>
                    );
                  })}
                  <td className="st st-first">{st.maxPossible}</td>
                  <td className={`st ${target !== null && stats.hasResults && st.total !== target ? 'bad' : ''}`}>{st.total}</td>
                  <td className="st strong">{st.cumLoad}</td>
                  {st.roleCounts.map((n, i) => (
                    <td key={i} className="st">
                      {n || ''}
                    </td>
                  ))}
                  <td className="st">{st.curLoad}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            {[
              { label: '감독 가능 교사 수', val: (i: number) => stats.slots[i].available, bad: (i: number) => stats.slots[i].available < stats.slots[i].need },
              { label: '필요 감독 수', val: (i: number) => stats.slots[i].need, bad: () => false },
              { label: '배정된 감독 수', val: (i: number) => stats.slots[i].assigned, bad: (i: number) => (stats.hasResults || stats.slots[i].assigned > 0) && stats.slots[i].assigned !== stats.slots[i].need },
            ].map((row, ri) => (
              <tr key={row.label} className={`f${ri}`}>
                <td className="sticky foot-label" colSpan={4}>
                  {row.label}
                </td>
                {slots.map((s, i) => (
                  <td key={s.key} className={`foot ${s.firstOfDay ? 'day-start' : ''} ${row.bad(i) ? 'bad' : ''}`}>
                    {row.val(i)}
                  </td>
                ))}
                <td className="foot st-first" colSpan={4 + p.roles.length} />
              </tr>
            ))}
          </tfoot>
        </table>
      </div>

      <details className="card issues-card" open={errCount > 0}>
        <summary>
          점검 결과 {errCount ? <span className="badge">{errCount}</span> : <span className="ok-text">이상 없음</span>}
        </summary>
        <IssueList issues={assignIssues} />
      </details>

      {editor && <CellEditor {...editor} onClose={() => setEditor(null)} />}
      {forb && <ConditionEditor {...forb} onClose={() => setForb(null)} />}

      {running && (
        <div className="modal-backdrop">
          <div className="running">
            <div className="spinner" />
            감독을 배정하는 중입니다…
          </div>
        </div>
      )}

      {pre && (
        <Modal
          title="배정할 수 없습니다"
          onClose={() => setPre(null)}
          width={600}
          footer={
            <>
              <div className="spacer" />
              <button className="btn primary" onClick={() => setPre(null)}>
                확인
              </button>
            </>
          }
        >
          <p>아래 항목을 먼저 해결해 주세요.</p>
          <IssueList issues={pre.map((msg) => ({ level: 'error', area: 'assign', msg }))} />
        </Modal>
      )}

      {needExtra && (
        <Modal
          title="추가 배정이 필요합니다"
          onClose={() => setNeedExtra(null)}
          width={520}
          footer={
            <>
              <div className="spacer" />
              <button className="btn" onClick={() => setNeedExtra(null)}>
                취소
              </button>
              <button
                className="btn primary"
                disabled={pick === null}
                onClick={() => {
                  const id = p.teachers[pick!].id;
                  const extra = { ...needExtra.extra, [id]: (needExtra.extra[id] || 0) + 1 };
                  setNeedExtra(null);
                  run(extra);
                }}
              >
                선택한 교사를 1시간 더 배정하고 계속
              </button>
            </>
          }
        >
          <p>
            <b>{slotLabel(slots[needExtra.slot])}</b>에 배정할 수 있는 교사가 없습니다. (다른 교사들은 배정할 시간을 모두 채웠거나 감독 불가입니다)
          </p>
          <p className="muted">아래 교사 중 한 명을 골라 배정할 시간보다 1시간 더 배정합니다. 현재 배정 시간이 적은 순입니다.</p>
          {needExtra.candidates.length ? (
            <div className="pick-list">
              {needExtra.candidates.map((c) => (
                <label key={c.t} className={pick === c.t ? 'on' : ''}>
                  <input type="radio" name="pick" checked={pick === c.t} onChange={() => setPick(c.t)} />
                  <b>{p.teachers[c.t].name}</b>
                  <span className="muted">
                    현재 {c.total}시간 / 배정할 시간 {c.target}
                  </span>
                </label>
              ))}
            </div>
          ) : (
            <p className="bad-text">추가로 배정할 수 있는 교사도 없습니다. 감독 불가(X) 표시나 필요 감독 수를 조정하세요.</p>
          )}
        </Modal>
      )}

      {result && (
        <Modal
          title="감독 배정 완료"
          onClose={() => setResult(null)}
          width={520}
          footer={
            <>
              <div className="spacer" />
              <button className="btn" onClick={() => setResult(null)}>
                닫기
              </button>
              <button
                className="btn accent"
                onClick={() => {
                  setResult(null);
                  setTab('chart');
                }}
              >
                감독표 보기 →
              </button>
            </>
          }
        >
          <div className="result-grid">
            <div className={result.metrics.conflicts ? 'bad' : 'good'}>
              <b>{result.metrics.conflicts}</b>
              <span>못 들어가는 고사실 충돌</span>
            </div>
            <div className={result.metrics.consec3 ? 'warn' : 'good'}>
              <b>{result.metrics.consec3}</b>
              <span>하루 3연속 이상 감독</span>
            </div>
            <div>
              <b>{result.metrics.consec2}</b>
              <span>하루 2연속 감독</span>
            </div>
            <div>
              <b>
                {result.metrics.loadMin}~{result.metrics.loadMax}
              </b>
              <span>누적 업무강도 범위</span>
            </div>
          </div>
          <p className="muted">
            {result.trials}회 시도 중 가장 좋은 결과 · {Math.round(result.ms)}ms · 업무강도 표준편차 {result.metrics.loadStd.toFixed(1)}
            {result.metrics.overTarget ? ` · 배정할 시간 초과 ${result.metrics.overTarget}명` : ''}
          </p>
          {result.metrics.conflicts > 0 && <p className="bad-text">빨간 테두리 칸이 못 들어가는 고사실에 배정된 칸입니다. 교환 기능으로 수동 조정해 주세요.</p>}
          {result.warnings.length > 0 && (
            <ul className="notes">
              {result.warnings.map((w, i) => (
                <li key={i}>{w}</li>
              ))}
            </ul>
          )}
          <p className="hint">결과가 마음에 들지 않으면 [감독 배정 시작]을 다시 누르세요. 매번 다른 조합을 시도합니다.</p>
        </Modal>
      )}
    </div>
  );
}

function CellEditor({ t, s, x, y, onClose }: { t: number; s: number; x: number; y: number; onClose: () => void }) {
  const { store, slots } = useApp();
  const p = store.project;
  const teacher = p.teachers[t];
  const slot = slots[s];
  const key = cellKey(teacher.id, slot.key);
  const c = p.cells[key];
  const [role, setRole] = useState(c?.role ?? p.roles[0]?.id ?? '');
  const [room, setRoom] = useState(c?.room ?? p.rooms[0]?.id ?? '');

  const put = (n: Cell | undefined) =>
    store.update((pp) => {
      const cells = { ...pp.cells };
      const v = n ? compact(n) : undefined;
      if (v) cells[key] = v;
      else delete cells[key];
      return { ...pp, cells };
    });

  const state = c?.x ? 'x' : c?.fixed ? 'fixed' : isAssigned(c) ? 'on' : 'empty';
  return (
    <Popover x={x} y={y} onClose={onClose} width={300}>
      <div className="pop-head">
        <b>{teacher.name}</b>
        <span className="muted">{slotLabel(slot)}</span>
      </div>
      <div className="pop-seg">
        <button className={state === 'empty' ? 'on' : ''} onClick={() => put(undefined)}>
          빈칸
        </button>
        <button className={state === 'x' ? 'on' : ''} onClick={() => put({ x: true })}>
          ✕ 불가
        </button>
        <button className={state === 'fixed' ? 'on' : ''} onClick={() => put({ ...c, x: false, fixed: true })}>
          📌 고정
        </button>
      </div>
      {c?.fixed && (
        <label className="pop-row">
          고정 보직
          <select className="inp" value={c.fixedRole ?? ''} onChange={(e) => put({ ...c, fixedRole: e.target.value || undefined })}>
            <option value="">자동(업무강도 기준)</option>
            {p.roles.map((r) => (
              <option key={r.id} value={r.id}>
                {r.name}
              </option>
            ))}
          </select>
        </label>
      )}
      {!c?.x && (
        <div className="pop-section">
          <div className="pop-sub">배정 결과 직접 지정</div>
          <div className="pop-row two">
            <select className="inp" value={role} onChange={(e) => setRole(e.target.value)} aria-label="보직">
              {p.roles.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.name}
                </option>
              ))}
            </select>
            <select className="inp" value={room} onChange={(e) => setRoom(e.target.value)} aria-label="고사실">
              {p.rooms.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.name}
                  {teacher.forbidden.includes(r.id) ? ' (불가)' : ''}
                </option>
              ))}
            </select>
          </div>
          <div className="pop-actions">
            <button className="btn sm" onClick={() => put({ fixed: c?.fixed, fixedRole: c?.fixedRole, on: false })} disabled={!c?.role}>
              결과 지우기
            </button>
            <button
              className="btn sm primary"
              onClick={() => {
                put({ ...c, on: true, role, room });
                onClose();
              }}
            >
              적용
            </button>
          </div>
        </div>
      )}
    </Popover>
  );
}

/** 교사별 근무 조건: 감독 가능 날짜 · 시수 고정 · 못 들어가는 고사실 */
function ConditionEditor({ t, x, y, onClose }: { t: number; x: number; y: number; onClose: () => void }) {
  const { store, slots, stats } = useApp();
  const p = store.project;
  const teacher = p.teachers[t];
  const st = stats.teachers[t];
  const days = useMemo(() => [...new Map(slots.map((s) => [s.dayId, s])).values()], [slots]);
  const allowed = teacher.availableDays ?? [];
  const limited = allowed.length > 0;
  const setTeacher = (fn: (tt: Teacher) => Teacher) =>
    store.update((pp) => ({ ...pp, teachers: pp.teachers.map((tt) => (tt.id === teacher.id ? fn(tt) : tt)) }));
  const toggleDay = (id: string) =>
    setTeacher((tt) => {
      const cur = tt.availableDays ?? [];
      const next = cur.includes(id) ? cur.filter((d) => d !== id) : [...cur, id];
      return { ...tt, availableDays: next.length ? next : undefined };
    });

  return (
    <Popover x={x} y={y} onClose={onClose} width={370}>
      <div className="pop-head">
        <b>{teacher.name}</b>
        <span className="muted">근무 조건</span>
      </div>

      <div className="pop-section first">
        <div className="pop-sub">감독 가능 날짜</div>
        <div className="pop-seg two">
          <button className={!limited ? 'on' : ''} onClick={() => setTeacher((tt) => ({ ...tt, availableDays: undefined }))}>
            모든 날 가능
          </button>
          <button
            className={limited ? 'on' : ''}
            onClick={() => setTeacher((tt) => ({ ...tt, availableDays: tt.availableDays?.length ? tt.availableDays : days[0] ? [days[0].dayId] : undefined }))}
          >
            선택한 날만
          </button>
        </div>
        {limited && (
          <div className="day-checks">
            {days.map((d) => (
              <label key={d.dayId} className={allowed.includes(d.dayId) ? 'on' : ''}>
                <input type="checkbox" checked={allowed.includes(d.dayId)} onChange={() => toggleDay(d.dayId)} />
                {fmtDay(d.date)}
              </label>
            ))}
          </div>
        )}
      </div>

      <div className="pop-section">
        <label className="pop-check">
          <input type="checkbox" checked={!!teacher.lockTarget} onChange={(e) => setTeacher((tt) => ({ ...tt, lockTarget: e.target.checked ? true : undefined }))} />
          배정할 시간 고정 <span className="muted">(자동 채우기에서 제외)</span>
        </label>
        <div className="pop-row two-inline">
          <span className="muted">배정할 시간</span>
          <NumberInput integer min={0} allowEmpty value={teacher.target} onChange={(v) => setTeacher((tt) => ({ ...tt, target: v }))} />
          <span className="muted">감독 가능 {st?.maxPossible ?? 0}시간</span>
        </div>
      </div>

      <div className="pop-section">
        <div className="pop-sub">못 들어가는 고사실</div>
        <div className="room-checks">
          {p.rooms.map((r) => (
            <label key={r.id} className={teacher.forbidden.includes(r.id) ? 'on' : ''}>
              <input
                type="checkbox"
                checked={teacher.forbidden.includes(r.id)}
                onChange={() =>
                  setTeacher((tt) => ({ ...tt, forbidden: tt.forbidden.includes(r.id) ? tt.forbidden.filter((f) => f !== r.id) : [...tt.forbidden, r.id] }))
                }
              />
              {r.name}
            </label>
          ))}
        </div>
      </div>

      <div className="pop-actions">
        <button className="btn sm" onClick={() => setTeacher((tt) => ({ ...tt, availableDays: undefined, forbidden: [], lockTarget: undefined }))}>
          조건 모두 지우기
        </button>
        <button className="btn sm primary" onClick={onClose}>
          닫기
        </button>
      </div>
    </Popover>
  );
}
