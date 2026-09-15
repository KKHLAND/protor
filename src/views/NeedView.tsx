import { Fragment, useMemo, useState, type ClipboardEvent, type KeyboardEvent } from 'react';
import { NumberInput, useApp } from '../ui';
import { fmtDay, needKey, roleColor, slotLabel } from '../model';
import type { Room } from '../types';

function roomGroups(rooms: Room[]) {
  const groups = new Map<string, string[]>();
  rooms.forEach((r) => {
    const m = /^(.+?)-/.exec(r.name);
    const g = m ? m[1] : '기타';
    if (!groups.has(g)) groups.set(g, []);
    groups.get(g)!.push(r.id);
  });
  return [...groups.entries()].map(([g, ids]) => ({ key: g, label: /^\d+$/.test(g) ? `${g}학년 (${g}-*)` : g, ids }));
}

export default function NeedView() {
  const { store, slots, stats, confirm, notify, setTab } = useApp();
  const p = store.project;
  const [fillRole, setFillRole] = useState(p.roles[0]?.id ?? '');
  const [fillVal, setFillVal] = useState<number | null>(1);
  const [fillDay, setFillDay] = useState('all');
  const [fillGroup, setFillGroup] = useState('all');
  const [copySrc, setCopySrc] = useState(slots[0]?.key ?? '');

  const groups = useMemo(() => roomGroups(p.rooms), [p.rooms]);
  const days = useMemo(() => [...new Map(slots.map((s) => [s.dayId, s])).values()], [slots]);
  const R = p.roles.length;
  const M = p.rooms.length;

  const setNeed = (updates: [string, number][], tag?: string) =>
    store.update(
      (pp) => {
        const need = { ...pp.need };
        for (const [k, v] of updates) {
          if (v > 0) need[k] = Math.floor(v);
          else delete need[k];
        }
        return { ...pp, need };
      },
      { tag },
    );

  const focusCell = (r: number, c: number) => {
    const el = document.querySelector<HTMLInputElement>(`input[data-nr="${r}"][data-nc="${c}"]`);
    if (el) {
      el.focus();
      el.select();
    }
  };

  const onKey = (e: KeyboardEvent<HTMLInputElement>, r: number, c: number) => {
    const el = e.currentTarget;
    if (e.key === 'ArrowDown' || e.key === 'Enter') {
      e.preventDefault();
      focusCell(r + 1, c);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusCell(r - 1, c);
    } else if (e.key === 'ArrowRight' && el.selectionEnd === el.value.length) {
      e.preventDefault();
      focusCell(r, c + 1);
    } else if (e.key === 'ArrowLeft' && el.selectionStart === 0) {
      e.preventDefault();
      focusCell(r, c - 1);
    }
  };

  const onPaste = (e: ClipboardEvent<HTMLInputElement>, r0: number, c0: number) => {
    const text = e.clipboardData.getData('text');
    if (!/[\t\n]/.test(text.trim())) return;
    e.preventDefault();
    const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
    const updates: [string, number][] = [];
    rows.forEach((cols, i) => {
      const ri = r0 + i;
      const s = slots[Math.floor(ri / R)];
      const role = p.roles[ri % R];
      if (!s || !role) return;
      cols.forEach((v, j) => {
        const room = p.rooms[c0 + j];
        if (!room) return;
        const n = Number(v.trim());
        updates.push([needKey(s.key, role.id, room.id), Number.isFinite(n) ? n : 0]);
      });
    });
    setNeed(updates);
    notify(`${updates.length}칸을 붙여넣었습니다.`, 'ok');
  };

  const quickFill = () => {
    if (!fillRole) return;
    const rooms = fillGroup === 'all' ? p.rooms.map((r) => r.id) : (groups.find((g) => g.key === fillGroup)?.ids ?? []);
    const updates: [string, number][] = [];
    slots.filter((s) => fillDay === 'all' || s.dayId === fillDay).forEach((s) => rooms.forEach((rid) => updates.push([needKey(s.key, fillRole, rid), fillVal ?? 0])));
    setNeed(updates);
    notify(`${updates.length}칸에 ${fillVal ?? 0}명을 입력했습니다.`, 'ok');
  };

  const copyToAll = () => {
    const src = slots.find((s) => s.key === copySrc);
    if (!src) return;
    const updates: [string, number][] = [];
    slots.forEach((s) => {
      if (s.key === src.key) return;
      p.roles.forEach((role) => p.rooms.forEach((room) => updates.push([needKey(s.key, role.id, room.id), p.need[needKey(src.key, role.id, room.id)] || 0])));
    });
    setNeed(updates);
    notify(`${slotLabel(src)} 입력을 모든 교시에 복사했습니다.`, 'ok');
  };

  if (!slots.length || !M || !R) {
    return (
      <div className="page">
        <div className="page-head">
          <div>
            <h2>필요 감독 수</h2>
          </div>
        </div>
        <div className="empty-state">
          <p>먼저 기본 정보에서 고사 일정, 고사실, 보직을 입력하세요.</p>
          <button className="btn primary" onClick={() => setTab('info')}>
            기본 정보로 이동
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h2>필요 감독 수</h2>
          <p className="sub">
            교시·보직별로 각 고사실에 필요한 감독 수를 입력합니다. 필요 없는 칸은 비워 두세요. 엑셀 표를 복사해 셀에 붙여넣을 수 있고, 방향키·Enter로
            이동합니다.
          </p>
        </div>
        <div className="head-stats">
          <span className="chip">총 필요 시간 {stats.totalNeed}</span>
          <button className="btn accent" onClick={() => setTab('assign')}>
            감독 배정으로 이동 →
          </button>
        </div>
      </div>

      <section className="card toolbar-card">
        <div className="toolbar">
          <span className="tool-label">빠른 채우기</span>
          <select className="inp" value={fillDay} onChange={(e) => setFillDay(e.target.value)} aria-label="날짜">
            <option value="all">모든 날짜</option>
            {days.map((d) => (
              <option key={d.dayId} value={d.dayId}>
                {fmtDay(d.date)}
              </option>
            ))}
          </select>
          <select className="inp" value={fillGroup} onChange={(e) => setFillGroup(e.target.value)} aria-label="고사실">
            <option value="all">모든 고사실</option>
            {groups.map((g) => (
              <option key={g.key} value={g.key}>
                {g.label}
              </option>
            ))}
          </select>
          <select className="inp" value={fillRole} onChange={(e) => setFillRole(e.target.value)} aria-label="보직">
            {p.roles.map((r) => (
              <option key={r.id} value={r.id}>
                {r.name}
              </option>
            ))}
          </select>
          <NumberInput integer min={0} value={fillVal} allowEmpty onChange={setFillVal} className="w60" />
          <span>명</span>
          <button className="btn primary sm" onClick={quickFill}>
            적용
          </button>
          <span className="sep" />
          <span className="tool-label">복사</span>
          <select className="inp" value={copySrc} onChange={(e) => setCopySrc(e.target.value)} aria-label="원본 교시">
            {slots.map((s) => (
              <option key={s.key} value={s.key}>
                {slotLabel(s)}
              </option>
            ))}
          </select>
          <button className="btn sm" onClick={copyToAll}>
            → 모든 교시에 복사
          </button>
          <span className="sep" />
          <button
            className="btn sm danger-outline"
            onClick={async () => {
              if (await confirm({ title: '필요 감독 수 비우기', message: '입력한 필요 감독 수를 모두 지웁니다.', ok: '모두 지우기', danger: true }))
                store.update((pp) => ({ ...pp, need: {} }));
            }}
          >
            모두 지우기
          </button>
        </div>
      </section>

      <div className="need-wrap">
        <table className="need">
          <thead>
            <tr>
              <th className="nc-date">날짜</th>
              <th className="nc-per">교시</th>
              <th className="nc-role">보직</th>
              {p.rooms.map((r) => (
                <th key={r.id} className="nc-room">
                  {r.name}
                </th>
              ))}
              <th>합계</th>
              <th className="nc-slot">교시 합계 / 가능 교사</th>
            </tr>
          </thead>
          <tbody>
            {slots.map((s, si) => {
              const ss = stats.slots[si];
              const dayFirst = s.firstOfDay;
              return (
                <Fragment key={s.key}>
                  {p.roles.map((role, ri) => {
                    const rowIdx = si * R + ri;
                    const rowSum = p.rooms.reduce((a, m) => a + (p.need[needKey(s.key, role.id, m.id)] || 0), 0);
                    return (
                      <tr key={role.id} className={`${ri === 0 ? 'per-first' : ''} ${ri === 0 && dayFirst ? 'day-first' : ''}`}>
                        {ri === 0 && (
                          <>
                            <td className="nc-date" rowSpan={R}>
                              {dayFirst || si === 0 ? fmtDay(s.date) : ''}
                            </td>
                            <td className="nc-per" rowSpan={R}>
                              {s.period}교시
                            </td>
                          </>
                        )}
                        <td className="nc-role" style={{ background: roleColor(ri) }}>
                          {role.name}
                        </td>
                        {p.rooms.map((room, mi) => {
                          const k = needKey(s.key, role.id, room.id);
                          const v = p.need[k] || 0;
                          return (
                            <td key={room.id} className={v ? 'has' : ''}>
                              <input
                                className="ncell"
                                data-nr={rowIdx}
                                data-nc={mi}
                                inputMode="numeric"
                                value={v || ''}
                                aria-label={`${slotLabel(s)} ${role.name} ${room.name}`}
                                onFocus={(e) => e.currentTarget.select()}
                                onKeyDown={(e) => onKey(e, rowIdx, mi)}
                                onPaste={(e) => onPaste(e, rowIdx, mi)}
                                onChange={(e) => {
                                  const n = Number(e.target.value.replace(/\D/g, ''));
                                  setNeed([[k, Number.isFinite(n) ? n : 0]], `need:${k}`);
                                }}
                              />
                            </td>
                          );
                        })}
                        <td className="sum">{rowSum || ''}</td>
                        {ri === 0 && (
                          <td className={`nc-slot ${ss.need > ss.available ? 'bad' : ''}`} rowSpan={R}>
                            <b>{ss.need}</b> / {ss.available}
                          </td>
                        )}
                      </tr>
                    );
                  })}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
