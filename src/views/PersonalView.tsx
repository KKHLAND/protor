import { useMemo, useState } from 'react';
import { useApp } from '../ui';
import { docTitle, fmtDay, personalRows, roleColor } from '../model';
import type { Project, Slot } from '../types';

export function PersonalCard({ project: p, slots, teacherId, hideEmpty, compact }: { project: Project; slots: Slot[]; teacherId: string; hideEmpty: boolean; compact?: boolean }) {
  const t = p.teachers.find((x) => x.id === teacherId);
  if (!t) return null;
  const rows = personalRows(p, slots, teacherId);
  const vis = hideEmpty ? rows.filter((r) => r.assigned) : rows;
  const total = rows.filter((r) => r.assigned).length;
  const counts = p.roles.map((r, i) => ({ name: r.name, n: rows.filter((x) => x.roleIndex === i).length })).filter((x) => x.n > 0);
  return (
    <div className={`pcard ${compact ? 'compact' : ''}`}>
      <div className="pcard-sub">{docTitle(p)}</div>
      <h3 className="pcard-title">{t.name} 선생님 감독 시간표</h3>
      <div className="pcard-sum">
        총 <b>{total}</b>시간{counts.length ? ' · ' + counts.map((c) => `${c.name} ${c.n}`).join(' · ') : ''}
      </div>
      <table className="ptable">
        <thead>
          <tr>
            <th>날짜</th>
            <th>교시</th>
            <th>고사실</th>
            <th>보직</th>
          </tr>
        </thead>
        <tbody>
          {vis.map((r, i) => {
            const first = i === 0 || vis[i - 1].slot.dayIdx !== r.slot.dayIdx;
            return (
              <tr key={r.slot.key} className={`${r.assigned ? 'assigned' : 'empty'} ${first ? 'day-first' : ''}`}>
                <td className="d">{first ? fmtDay(r.slot.date) : ''}</td>
                <td>{r.slot.period}교시</td>
                <td className="room">{r.roomName}</td>
                <td>{r.assigned && <span className="role-tag" style={{ background: roleColor(r.roleIndex) }}>{r.roleName}</span>}</td>
              </tr>
            );
          })}
          {!vis.length && (
            <tr>
              <td colSpan={4} className="empty-row">
                배정된 감독이 없습니다.
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

export default function PersonalView() {
  const { store, slots, stats, print, setTab } = useApp();
  const p = store.project;
  const [q, setQ] = useState('');
  const [selId, setSelId] = useState(p.teachers[0]?.id ?? '');
  const [hideEmpty, setHideEmpty] = useState(false);
  const idx = Math.max(0, p.teachers.findIndex((t) => t.id === selId));
  const current = p.teachers[idx];

  const list = useMemo(() => p.teachers.map((t, i) => ({ t, i })).filter(({ t }) => !q.trim() || t.name.includes(q.trim())), [p.teachers, q]);

  if (!p.teachers.length) {
    return (
      <div className="page">
        <div className="empty-state">
          <p>감독 교사가 없습니다.</p>
          <button className="btn primary" onClick={() => setTab('info')}>
            기본 정보로 이동
          </button>
        </div>
      </div>
    );
  }

  const allIds = p.teachers.map((t) => t.id);
  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h2>개인 시간표</h2>
          <p className="sub">교사별 감독 시간표를 확인하고, 한 명씩 또는 전체 교사를 한 번에 인쇄합니다. 인쇄 창에서 “PDF로 저장”을 고르면 PDF로 저장됩니다.</p>
        </div>
      </div>
      {!stats.hasResults && (
        <div className="banner">
          아직 감독 배정 결과가 없습니다.
          <button className="btn sm primary" onClick={() => setTab('assign')}>
            감독 배정으로 이동
          </button>
        </div>
      )}
      <section className="card toolbar-card">
        <div className="toolbar">
          <label className="inline">
            <input type="checkbox" checked={hideEmpty} onChange={(e) => setHideEmpty(e.target.checked)} /> 배정된 시간만 표시
          </label>
          <div className="spacer" />
          <button className="btn sm" onClick={() => current && print({ kind: 'personal', teacherIds: [current.id], layout: 'page', hideEmpty })}>
            이 선생님 인쇄
          </button>
          <button className="btn sm primary" onClick={() => print({ kind: 'personal', teacherIds: allIds, layout: 'page', hideEmpty })}>
            전체 교사 인쇄 (1인 1쪽)
          </button>
          <button className="btn sm" onClick={() => print({ kind: 'personal', teacherIds: allIds, layout: 'compact', hideEmpty: true })}>
            전체 교사 모아 인쇄 (한 쪽에 4명)
          </button>
        </div>
      </section>
      <div className="personal-layout">
        <aside className="card teacher-list">
          <input className="inp" placeholder="이름 검색" value={q} onChange={(e) => setQ(e.target.value)} />
          <ul>
            {list.map(({ t, i }) => (
              <li key={t.id}>
                <button className={t.id === current?.id ? 'on' : ''} onClick={() => setSelId(t.id)}>
                  <span className="muted">{i + 1}</span> {t.name}
                  <span className="pill">{stats.teachers[i]?.total ?? 0}시간</span>
                </button>
              </li>
            ))}
          </ul>
        </aside>
        <div className="personal-main">
          <div className="pnav">
            <button className="btn sm" disabled={idx <= 0} onClick={() => setSelId(p.teachers[idx - 1].id)}>
              ← 이전
            </button>
            <span className="muted">
              {idx + 1} / {p.teachers.length}
            </span>
            <button className="btn sm" disabled={idx >= p.teachers.length - 1} onClick={() => setSelId(p.teachers[idx + 1].id)}>
              다음 →
            </button>
          </div>
          <div className="doc pdoc">{current && <PersonalCard project={p} slots={slots} teacherId={current.id} hideEmpty={hideEmpty} />}</div>
        </div>
      </div>
    </div>
  );
}
