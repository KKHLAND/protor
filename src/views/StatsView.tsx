import { useMemo, useState } from 'react';
import { useApp } from '../ui';
import { fmtDay, nextExamProject, roleColor, type Stats } from '../model';
import type { Project, Slot } from '../types';

type SortKey = 'no' | 'name' | 'total' | 'cum' | 'cur' | 'consec';

export function StatsTable({ project: p, slots, stats, sort = 'no', onSort }: { project: Project; slots: Slot[]; stats: Stats; sort?: SortKey; onSort?: (k: SortKey) => void }) {
  const days = [...new Map(slots.map((s) => [s.dayIdx, s])).values()];
  const maxLoad = Math.max(1, ...stats.teachers.map((s) => s.cumLoad));
  const rows = p.teachers.map((t, i) => ({ t, i, st: stats.teachers[i] }));
  const cmp: Record<SortKey, (a: (typeof rows)[0], b: (typeof rows)[0]) => number> = {
    no: (a, b) => a.i - b.i,
    name: (a, b) => a.t.name.localeCompare(b.t.name, 'ko'),
    total: (a, b) => b.st.total - a.st.total,
    cum: (a, b) => b.st.cumLoad - a.st.cumLoad,
    cur: (a, b) => b.st.curLoad - a.st.curLoad,
    consec: (a, b) => b.st.maxConsec - a.st.maxConsec,
  };
  rows.sort(cmp[sort]);
  const H = ({ k, children }: { k: SortKey; children: React.ReactNode }) => (
    <th className={onSort ? 'clickable' : ''} onClick={() => onSort?.(k)}>
      {children}
      {onSort && sort === k ? ' ▾' : ''}
    </th>
  );
  return (
    <table className="tbl stats">
      <thead>
        <tr>
          <H k="no">순번</H>
          <H k="name">이름</H>
          <th>배정할 시간</th>
          <H k="total">총 감독</H>
          {days.map((d) => (
            <th key={d.dayIdx}>{fmtDay(d.date)}</th>
          ))}
          <H k="consec">최대 연속</H>
          {p.roles.map((r, i) => (
            <th key={r.id} style={{ boxShadow: `inset 0 -3px 0 ${roleColor(i)}` }}>
              {r.name}
            </th>
          ))}
          <H k="cur">현재 업무강도</H>
          <th>이전 누적</th>
          <H k="cum">누적 업무강도</H>
        </tr>
      </thead>
      <tbody>
        {rows.map(({ t, i, st }) => (
          <tr key={t.id}>
            <td className="muted">{i + 1}</td>
            <td className="name">{t.name}</td>
            <td>{t.target ?? '—'}</td>
            <td className={t.target !== null && st.total !== t.target && stats.hasResults ? 'bad' : ''}>{st.total}</td>
            {days.map((d) => (
              <td key={d.dayIdx}>{st.dayCounts[d.dayIdx] || ''}</td>
            ))}
            <td className={st.maxConsec >= 3 ? 'bad' : ''}>{st.maxConsec || ''}</td>
            {st.roleCounts.map((n, k) => (
              <td key={k}>{n || ''}</td>
            ))}
            <td>{st.curLoad}</td>
            <td className="muted">{t.prevLoad || 0}</td>
            <td className="bar-cell">
              <span className="bar">
                <i style={{ width: `${(st.cumLoad / maxLoad) * 100}%` }} />
              </span>
              <b>{st.cumLoad}</b>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export default function StatsView() {
  const { store, slots, stats, print, confirm, notify, saveJson, setTab } = useApp();
  const p = store.project;
  const [sort, setSort] = useState<SortKey>('no');

  const summary = useMemo(() => {
    const loads = stats.teachers.map((s) => s.cumLoad);
    const cur = stats.teachers.map((s) => s.curLoad);
    const n = Math.max(1, loads.length);
    const mean = loads.reduce((a, b) => a + b, 0) / n;
    const std = Math.sqrt(loads.reduce((a, b) => a + (b - mean) ** 2, 0) / n);
    const totals = stats.teachers.map((s) => s.total);
    return {
      totalHours: totals.reduce((a, b) => a + b, 0),
      avgHours: totals.reduce((a, b) => a + b, 0) / n,
      minH: totals.length ? Math.min(...totals) : 0,
      maxH: totals.length ? Math.max(...totals) : 0,
      mean,
      std,
      min: loads.length ? Math.min(...loads) : 0,
      max: loads.length ? Math.max(...loads) : 0,
      curMin: cur.length ? Math.min(...cur) : 0,
      curMax: cur.length ? Math.max(...cur) : 0,
      consec3: stats.teachers.filter((s) => s.maxConsec >= 3).length,
      conflicts: stats.teachers.reduce((a, s) => a + s.conflicts, 0),
    };
  }, [stats]);

  const nextExam = async () => {
    const ok = await confirm({
      title: '다음 고사 준비',
      message: (
        <>
          <p>
            교사별 <b>누적 업무강도</b>를 “이전 누적 업무강도”로 옮기고, 고사 일정·필요 감독 수·배정 결과를 비웁니다. 교사·고사실·보직 목록은 그대로
            남습니다.
          </p>
          <p className="muted">현재 고사 결과를 보관하려면 먼저 [프로젝트 저장]을 누르세요.</p>
        </>
      ),
      ok: '다음 고사 준비',
      danger: true,
    });
    if (!ok) return;
    store.update((pp) => nextExamProject(pp, stats));
    notify('누적 업무강도를 이월했습니다. 새 고사 일정을 입력하세요.', 'ok');
    setTab('info');
  };

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h2>감독 현황</h2>
          <p className="sub">교사별 감독 시간, 날짜별 분포, 보직 수, 업무강도를 확인합니다. 머리글을 누르면 정렬됩니다.</p>
        </div>
      </div>
      <div className="stat-cards">
        <div className="stat-card">
          <span>총 감독 시간</span>
          <b>{summary.totalHours}</b>
          <small>
            1인 평균 {summary.avgHours.toFixed(1)}시간 ({summary.minH}~{summary.maxH})
          </small>
        </div>
        <div className="stat-card">
          <span>이번 고사 업무강도</span>
          <b>
            {summary.curMin}~{summary.curMax}
          </b>
          <small>교사별 최소~최대</small>
        </div>
        <div className="stat-card">
          <span>누적 업무강도</span>
          <b>
            {summary.min}~{summary.max}
          </b>
          <small>
            평균 {summary.mean.toFixed(0)} · 표준편차 {summary.std.toFixed(1)}
          </small>
        </div>
        <div className={`stat-card ${summary.consec3 ? 'warn' : ''}`}>
          <span>하루 3연속 이상</span>
          <b>{summary.consec3}명</b>
          <small>같은 날 연속 감독</small>
        </div>
        <div className={`stat-card ${summary.conflicts ? 'bad' : ''}`}>
          <span>고사실 충돌</span>
          <b>{summary.conflicts}건</b>
          <small>못 들어가는 고사실 배정</small>
        </div>
      </div>
      <section className="card toolbar-card">
        <div className="toolbar">
          <button className="btn sm" onClick={() => print({ kind: 'stats' })}>
            인쇄
          </button>
          <button className="btn sm" onClick={saveJson}>
            프로젝트 저장
          </button>
          <div className="spacer" />
          <button className="btn sm danger-outline" onClick={nextExam} disabled={!stats.hasResults}>
            다음 고사 준비 (누적 업무강도 이월)
          </button>
        </div>
      </section>
      <div className="card tbl-wrap">
        <StatsTable project={p} slots={slots} stats={stats} sort={sort} onSort={setSort} />
      </div>
    </div>
  );
}
