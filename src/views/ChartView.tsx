import { useMemo, useState, type CSSProperties } from 'react';
import { Seg, useApp } from '../ui';
import DailyShareCard from './DailyShare';
import { cellKey, docTitle, fmtDay, isAssigned, needKey, roleColor } from '../model';
import type { Project, Slot } from '../types';

export function ChartTable({
  project: p,
  slots,
  dayIdx,
  display,
  highlightId,
  hideEmpty,
}: {
  project: Project;
  slots: Slot[];
  dayIdx?: number;
  display: 'name' | 'number';
  highlightId?: string;
  hideEmpty: boolean;
}) {
  const buckets = useMemo(() => {
    const map = new Map<string, { id: string; label: string }[]>();
    p.teachers.forEach((t, ti) =>
      slots.forEach((s) => {
        const c = p.cells[cellKey(t.id, s.key)];
        if (!isAssigned(c) || !c!.role || !c!.room) return;
        const k = `${s.key}|${c!.role}|${c!.room}`;
        if (!map.has(k)) map.set(k, []);
        map.get(k)!.push({ id: t.id, label: display === 'name' ? t.name : String(ti + 1) });
      }),
    );
    return map;
  }, [p, slots, display]);

  const vis = slots.filter((s) => dayIdx === undefined || s.dayIdx === dayIdx);
  const rolesFor = (s: Slot) => {
    const all = p.roles.map((r, ri) => ({ r, ri }));
    if (!hideEmpty) return all;
    const v = all.filter(({ r }) => p.rooms.some((m) => (p.need[needKey(s.key, r.id, m.id)] || 0) > 0));
    return v.length ? v : all;
  };
  const dayRows = new Map<number, number>();
  vis.forEach((s) => dayRows.set(s.dayIdx, (dayRows.get(s.dayIdx) || 0) + rolesFor(s).length));

  return (
    <table className="chart" style={{ '--rooms': p.rooms.length } as CSSProperties}>
      <thead>
        <tr>
          <th className="c-date">날짜</th>
          <th className="c-per">교시</th>
          <th className="c-role">보직</th>
          {p.rooms.map((m) => (
            <th key={m.id}>{m.name}</th>
          ))}
          <th className="c-sum">합계</th>
        </tr>
      </thead>
      <tbody>
        {vis.map((s, i) => {
          const roles = rolesFor(s);
          const dayFirst = i === 0 || vis[i - 1].dayIdx !== s.dayIdx;
          return roles.map(({ r, ri }, k) => {
            const rowNeed = p.rooms.reduce((a, m) => a + (p.need[needKey(s.key, r.id, m.id)] || 0), 0);
            return (
              <tr key={`${s.key}|${r.id}`} className={`${k === 0 ? 'per-first' : ''} ${k === 0 && dayFirst ? 'day-first' : ''}`}>
                {k === 0 && dayFirst && (
                  <td className="c-date" rowSpan={dayRows.get(s.dayIdx)}>
                    {fmtDay(s.date)}
                  </td>
                )}
                {k === 0 && (
                  <td className="c-per" rowSpan={roles.length}>
                    {s.period}교시
                  </td>
                )}
                <td className="c-role" style={{ background: roleColor(ri) }}>
                  {r.name}
                </td>
                {p.rooms.map((m) => {
                  const list = buckets.get(`${s.key}|${r.id}|${m.id}`) || [];
                  const need = p.need[needKey(s.key, r.id, m.id)] || 0;
                  const cls = ['names', need === 0 && !list.length ? 'noneed' : '', list.length !== need ? 'mismatch' : ''].join(' ');
                  return (
                    <td key={m.id} className={cls}>
                      {list.map((x) => (
                        <div key={x.id} className={x.id === highlightId ? 'me' : ''}>
                          {x.label}
                        </div>
                      ))}
                    </td>
                  );
                })}
                <td className="c-sum">{rowNeed || ''}</td>
              </tr>
            );
          });
        })}
      </tbody>
    </table>
  );
}

export default function ChartView() {
  const { store, slots, stats, print, exportExcel, setTab } = useApp();
  const p = store.project;
  const [display, setDisplay] = useState<'name' | 'number'>('name');
  const [hl, setHl] = useState('');
  const [mode, setMode] = useState<'all' | 'daily'>('all');
  const [day, setDay] = useState(0);
  const [hideEmpty, setHideEmpty] = useState(true);
  const days = useMemo(() => [...new Map(slots.map((s) => [s.dayIdx, s])).values()], [slots]);
  const curDay = days.find((d) => d.dayIdx === day) ?? days[0];

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h2>감독표</h2>
          <p className="sub">고사실별 감독 교사를 확인하고 인쇄합니다. 인쇄 창에서 “PDF로 저장”을 고르면 PDF 파일로 저장됩니다.</p>
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

      <DailyShareCard />

      <section className="card toolbar-card">
        <div className="toolbar">
          <Seg
            value={mode}
            onChange={setMode}
            options={[
              { v: 'all', label: '전체' },
              { v: 'daily', label: '일별' },
            ]}
          />
          {mode === 'daily' && (
            <div className="seg">
              {days.map((d) => (
                <button key={d.dayIdx} className={curDay?.dayIdx === d.dayIdx ? 'on' : ''} onClick={() => setDay(d.dayIdx)}>
                  {fmtDay(d.date)}
                </button>
              ))}
            </div>
          )}
          <span className="sep" />
          <Seg
            value={display}
            onChange={setDisplay}
            options={[
              { v: 'name', label: '이름으로 표시' },
              { v: 'number', label: '감독번호로 표시' },
            ]}
          />
          <select className="inp" value={hl} onChange={(e) => setHl(e.target.value)} aria-label="강조할 교사">
            <option value="">교사 강조 표시 없음</option>
            {p.teachers.map((t, i) => (
              <option key={t.id} value={t.id}>
                {i + 1}. {t.name}
              </option>
            ))}
          </select>
          <label className="inline">
            <input type="checkbox" checked={hideEmpty} onChange={(e) => setHideEmpty(e.target.checked)} /> 필요 없는 보직 줄 숨기기
          </label>
          <div className="spacer" />
          <button className="btn sm" onClick={() => print({ kind: 'chart', mode: 'all', display, highlightId: hl || undefined, hideEmpty })}>
            전체 인쇄
          </button>
          <button className="btn sm" onClick={() => print({ kind: 'chart', mode: 'daily', display, highlightId: hl || undefined, hideEmpty })}>
            일별 인쇄 (날짜별 한 쪽)
          </button>
          {mode === 'daily' && curDay && (
            <button className="btn sm" onClick={() => print({ kind: 'chart', mode: 'daily', dayIdx: curDay.dayIdx, display, highlightId: hl || undefined, hideEmpty })}>
              {fmtDay(curDay.date)}만 인쇄
            </button>
          )}
          <button className="btn sm accent" onClick={exportExcel}>
            엑셀
          </button>
        </div>
      </section>

      <div className="doc">
        <div className="doc-title">
          {docTitle(p) || '고사'} 감독표{mode === 'daily' && curDay ? ` — ${fmtDay(curDay.date)}` : ''}
        </div>
        {display === 'number' && (
          <div className="number-key">
            {p.teachers.map((t, i) => (
              <span key={t.id}>
                <b>{i + 1}</b> {t.name}
              </span>
            ))}
          </div>
        )}
        <div className="chart-scroll">
          <ChartTable project={p} slots={slots} dayIdx={mode === 'daily' ? curDay?.dayIdx : undefined} display={display} highlightId={hl || undefined} hideEmpty={hideEmpty} />
        </div>
      </div>
    </div>
  );
}
