import { useState } from 'react';
import { BulkPaste, Modal, NumberInput, useApp } from '../ui';
import { fmtDay, parseDate, roleColor, uid } from '../model';
import type { Day, Project, Role, Room, Teacher } from '../types';

const pad = (n: number) => String(n).padStart(2, '0');
const toISO = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

function normalizeDate(s: string, year: number): string {
  const t = s.trim();
  let m = /^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/.exec(t);
  if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
  m = /^(\d{1,2})\s*[-./월]\s*(\d{1,2})/.exec(t);
  if (m) return `${year}-${pad(+m[1])}-${pad(+m[2])}`;
  return '';
}

type ListKey = 'days' | 'rooms' | 'teachers' | 'roles';

function move<T>(list: T[], i: number, dir: -1 | 1): T[] {
  const j = i + dir;
  if (j < 0 || j >= list.length) return list;
  const next = [...list];
  [next[i], next[j]] = [next[j], next[i]];
  return next;
}

export default function InfoView() {
  const { store, slots, stats, notify, openFile, downloadTemplate } = useApp();
  const p = store.project;
  const [paste, setPaste] = useState<ListKey | null>(null);
  const [gen, setGen] = useState(false);
  const [q, setQ] = useState('');

  const year = Number(p.days.find((d) => d.date)?.date.slice(0, 4)) || new Date().getFullYear();

  const edit = <K extends ListKey>(key: K, fn: (list: Project[K]) => Project[K], tag?: string) =>
    store.update((pp) => ({ ...pp, [key]: fn(pp[key]) }), { tag });
  const remove = <K extends ListKey>(key: K, id: string) =>
    store.update((pp) => ({ ...pp, [key]: (pp[key] as { id: string }[]).filter((x) => x.id !== id) }), { clean: true });

  const addDay = () =>
    edit('days', (days) => {
      const last = [...days].sort((a, b) => a.date.localeCompare(b.date)).pop();
      const d = last && parseDate(last.date);
      let date = '';
      if (d) {
        d.setDate(d.getDate() + (d.getDay() === 5 ? 3 : 1));
        date = toISO(d);
      }
      return [...days, { id: uid(), date, start: 1, end: last?.end ?? 3 }];
    });

  const applyPaste = (key: ListKey, lines: string[], mode: 'append' | 'replace') => {
    const cols = (l: string) => l.split(/\t|,/).map((s) => s.trim());
    let added = 0;
    store.update(
      (pp) => {
        const next: Project = { ...pp };
        if (key === 'days') {
          const items: Day[] = [];
          lines.forEach((l) => {
            const c = l.split(/[\t,]+|\s{1,}(?=\d+\s*$)|\s+(?=\d+\s+\d+\s*$)/).map((s) => s.trim()).filter(Boolean);
            const date = normalizeDate(c[0] ?? '', year);
            if (!date) return;
            const start = Number(c[1]) || 1;
            const end = Number(c[2]) || start;
            items.push({ id: uid(), date, start, end });
          });
          added = items.length;
          next.days = mode === 'replace' ? items : [...pp.days, ...items];
        } else if (key === 'rooms') {
          const items: Room[] = lines.flatMap((l) => l.split('\t')).map((s) => s.trim()).filter(Boolean).map((name) => ({ id: uid(), name }));
          added = items.length;
          next.rooms = mode === 'replace' ? items : [...pp.rooms, ...items];
        } else if (key === 'teachers') {
          const items: Teacher[] = lines
            .map(cols)
            .filter((c) => c[0])
            .map((c) => ({ id: uid(), name: c[0], prevLoad: Number(c[1]) || 0, target: null, forbidden: [] }));
          added = items.length;
          next.teachers = mode === 'replace' ? items : [...pp.teachers, ...items];
        } else {
          const items: Role[] = lines
            .map(cols)
            .filter((c) => c[0])
            .map((c) => ({ id: uid(), name: c[0], weight: Number(c[1]) || 0 }));
          added = items.length;
          next.roles = mode === 'replace' ? items : [...pp.roles, ...items];
        }
        return next;
      },
      { clean: mode === 'replace' },
    );
    setTimeout(() => notify(added ? `${added}개 항목을 넣었습니다.` : '인식된 항목이 없습니다. 형식을 확인해 주세요.', added ? 'ok' : 'warn'));
  };

  const filteredTeachers = p.teachers.map((t, i) => ({ t, i })).filter(({ t }) => !q.trim() || t.name.includes(q.trim()));

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h2>기본 정보</h2>
          <p className="sub">고사 일정, 고사실, 감독 교사, 보직을 입력합니다. 엑셀에서 복사한 목록은 [붙여넣기]로 한 번에 넣을 수 있습니다.</p>
        </div>
        <div className="head-stats">
          <span className="chip">시험 시간 {slots.length}</span>
          <span className="chip">고사실 {p.rooms.length}</span>
          <span className="chip">교사 {p.teachers.length}</span>
        </div>
      </div>

      <section className="card excel-card">
        <div className="excel-card-main">
          <h3>엑셀로 한 번에 입력하기</h3>
          <p>
            양식을 내려받아 엑셀에서 채운 뒤 올리면 일정·고사실·교사·보직이 한꺼번에 입력됩니다. 기본 정보를 올린 다음 양식을 다시 받으면 교시·보직 줄과
            교사 명단이 채워져 있어, <b>필요 감독 수</b>와 <b>감독 불가(x)·고정(1)</b> 표시까지 엑셀에서 작성할 수 있습니다.
          </p>
        </div>
        <div className="excel-card-actions">
          <button className="btn primary" onClick={downloadTemplate}>
            ① 엑셀 양식 내려받기
          </button>
          <button className="btn" onClick={openFile}>
            ② 작성한 엑셀 올리기
          </button>
        </div>
      </section>

      <section className="card">
        <div className="card-head">
          <h3>고사 이름</h3>
        </div>
        <div className="meta-row">
          <label>
            학교명
            <input
              className="inp"
              value={p.meta.school}
              placeholder="예: OO고등학교"
              onChange={(e) => store.update((pp) => ({ ...pp, meta: { ...pp.meta, school: e.target.value } }), { tag: 'meta:school' })}
            />
          </label>
          <label className="grow">
            고사 이름
            <input
              className="inp"
              value={p.meta.title}
              placeholder="예: 2026학년도 2학기 1회고사"
              onChange={(e) => store.update((pp) => ({ ...pp, meta: { ...pp.meta, title: e.target.value } }), { tag: 'meta:title' })}
            />
          </label>
        </div>
      </section>

      <div className="grid-2">
        <section className="card">
          <div className="card-head">
            <h3>
              1. 고사 날짜 및 교시 <span className="count">{p.days.length}일</span>
            </h3>
            <div className="card-actions">
              <button className="btn sm" onClick={() => setPaste('days')}>
                붙여넣기
              </button>
              <button className="btn sm primary" onClick={addDay}>
                + 날짜 추가
              </button>
            </div>
          </div>
          <table className="tbl">
            <thead>
              <tr>
                <th>순번</th>
                <th>날짜</th>
                <th>시작교시</th>
                <th>종료교시</th>
                <th>시간</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {p.days.map((d, i) => (
                <tr key={d.id}>
                  <td className="muted">{i + 1}</td>
                  <td>
                    <input
                      type="date"
                      className="inp"
                      value={d.date}
                      onChange={(e) => edit('days', (ds) => ds.map((x) => (x.id === d.id ? { ...x, date: e.target.value } : x)), `day:${d.id}`)}
                    />
                    <span className="weekday">{d.date ? fmtDay(d.date).slice(-3) : ''}</span>
                  </td>
                  <td>
                    <NumberInput integer min={1} max={12} value={d.start} onChange={(v) => edit('days', (ds) => ds.map((x) => (x.id === d.id ? { ...x, start: v ?? 1 } : x)), `day:${d.id}:s`)} />
                  </td>
                  <td>
                    <NumberInput integer min={1} max={12} value={d.end} onChange={(v) => edit('days', (ds) => ds.map((x) => (x.id === d.id ? { ...x, end: v ?? 1 } : x)), `day:${d.id}:e`)} />
                  </td>
                  <td className="muted">{Math.max(0, d.end - d.start + 1)}</td>
                  <td>
                    <button className="icon-btn" title="삭제" onClick={() => remove('days', d.id)}>
                      🗑
                    </button>
                  </td>
                </tr>
              ))}
              {!p.days.length && (
                <tr>
                  <td colSpan={6} className="empty-row">
                    [+ 날짜 추가]로 고사 날짜를 입력하세요.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
          <p className="hint">날짜순으로 자동 정렬되어 표에 나타납니다. 총 시험 시간: {slots.length}시간</p>
        </section>

        <section className="card">
          <div className="card-head">
            <h3>
              4. 보직과 업무강도 <span className="count">{p.roles.length}개</span>
            </h3>
            <div className="card-actions">
              <button className="btn sm" onClick={() => setPaste('roles')}>
                붙여넣기
              </button>
              <button className="btn sm primary" onClick={() => edit('roles', (rs) => [...rs, { id: uid(), name: '', weight: 0 }])}>
                + 보직 추가
              </button>
            </div>
          </div>
          <table className="tbl">
            <thead>
              <tr>
                <th>순서</th>
                <th>보직</th>
                <th>업무강도</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {p.roles.map((r, i) => (
                <tr key={r.id}>
                  <td>
                    <span className="swatch" style={{ background: roleColor(i) }} />
                    {i + 1}
                  </td>
                  <td>
                    <input className="inp" value={r.name} placeholder="예: 정감독" onChange={(e) => edit('roles', (rs) => rs.map((x) => (x.id === r.id ? { ...x, name: e.target.value } : x)), `role:${r.id}`)} />
                  </td>
                  <td>
                    <NumberInput min={0} value={r.weight} onChange={(v) => edit('roles', (rs) => rs.map((x) => (x.id === r.id ? { ...x, weight: v ?? 0 } : x)), `role:${r.id}:w`)} />
                  </td>
                  <td className="row-actions">
                    <button className="icon-btn" title="위로" onClick={() => edit('roles', (rs) => move(rs, i, -1))}>
                      ▲
                    </button>
                    <button className="icon-btn" title="아래로" onClick={() => edit('roles', (rs) => move(rs, i, 1))}>
                      ▼
                    </button>
                    <button className="icon-btn" title="삭제" onClick={() => remove('roles', r.id)}>
                      🗑
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="hint">
            업무강도는 보직의 부담 정도입니다(예: 정감독 100, 부감독 50). 위에 있는 보직부터, 누적 업무강도가 낮은 교사에게 먼저 배정되고 마지막에 교사 간
            업무강도가 고르게 되도록 조정됩니다.
          </p>
        </section>
      </div>

      <div className="grid-2">
        <section className="card">
          <div className="card-head">
            <h3>
              2. 고사실 <span className="count">{p.rooms.length}개</span>
            </h3>
            <div className="card-actions">
              <button className="btn sm" onClick={() => setGen(true)}>
                학년·반 생성
              </button>
              <button className="btn sm" onClick={() => setPaste('rooms')}>
                붙여넣기
              </button>
              <button className="btn sm primary" onClick={() => edit('rooms', (rs) => [...rs, { id: uid(), name: '' }])}>
                + 추가
              </button>
            </div>
          </div>
          <div className="scroll-list">
            <table className="tbl">
              <thead>
                <tr>
                  <th>순번</th>
                  <th>고사실명</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {p.rooms.map((r, i) => (
                  <tr key={r.id}>
                    <td className="muted">{i + 1}</td>
                    <td>
                      <input className="inp" value={r.name} placeholder="예: 1-1" onChange={(e) => edit('rooms', (rs) => rs.map((x) => (x.id === r.id ? { ...x, name: e.target.value } : x)), `room:${r.id}`)} />
                    </td>
                    <td className="row-actions">
                      <button className="icon-btn" title="위로" onClick={() => edit('rooms', (rs) => move(rs, i, -1))}>
                        ▲
                      </button>
                      <button className="icon-btn" title="아래로" onClick={() => edit('rooms', (rs) => move(rs, i, 1))}>
                        ▼
                      </button>
                      <button className="icon-btn" title="삭제" onClick={() => remove('rooms', r.id)}>
                        🗑
                      </button>
                    </td>
                  </tr>
                ))}
                {!p.rooms.length && (
                  <tr>
                    <td colSpan={3} className="empty-row">
                      [학년·반 생성]으로 1-1, 1-2 … 형식의 고사실을 한 번에 만들 수 있습니다.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>

        <section className="card">
          <div className="card-head">
            <h3>
              3. 감독 교사 <span className="count">{p.teachers.length}명</span>
            </h3>
            <div className="card-actions">
              <input className="inp search" placeholder="이름 검색" value={q} onChange={(e) => setQ(e.target.value)} />
              <button
                className="btn sm"
                onClick={() => edit('teachers', (ts) => [...ts].sort((a, b) => a.name.localeCompare(b.name, 'ko')))}
                title="가나다순 정렬"
              >
                가나다순
              </button>
              <button className="btn sm" onClick={() => setPaste('teachers')}>
                붙여넣기
              </button>
              <button className="btn sm primary" onClick={() => edit('teachers', (ts) => [...ts, { id: uid(), name: '', prevLoad: 0, target: null, forbidden: [] }])}>
                + 추가
              </button>
            </div>
          </div>
          <div className="scroll-list">
            <table className="tbl">
              <thead>
                <tr>
                  <th>순번</th>
                  <th>이름</th>
                  <th title="이전 시험까지의 누적 업무강도(선택 입력)">이전 누적 업무강도</th>
                  <th>이번 누적</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {filteredTeachers.map(({ t, i }) => (
                  <tr key={t.id}>
                    <td className="muted">{i + 1}</td>
                    <td>
                      <input className="inp" value={t.name} placeholder="이름" onChange={(e) => edit('teachers', (ts) => ts.map((x) => (x.id === t.id ? { ...x, name: e.target.value } : x)), `t:${t.id}`)} />
                    </td>
                    <td>
                      <NumberInput min={0} value={t.prevLoad} onChange={(v) => edit('teachers', (ts) => ts.map((x) => (x.id === t.id ? { ...x, prevLoad: v ?? 0 } : x)), `t:${t.id}:l`)} />
                    </td>
                    <td className="muted">{stats.hasResults ? stats.teachers[i]?.cumLoad : '—'}</td>
                    <td className="row-actions">
                      <button className="icon-btn" title="위로" onClick={() => edit('teachers', (ts) => move(ts, i, -1))}>
                        ▲
                      </button>
                      <button className="icon-btn" title="아래로" onClick={() => edit('teachers', (ts) => move(ts, i, 1))}>
                        ▼
                      </button>
                      <button className="icon-btn" title="삭제" onClick={() => remove('teachers', t.id)}>
                        🗑
                      </button>
                    </td>
                  </tr>
                ))}
                {!p.teachers.length && (
                  <tr>
                    <td colSpan={5} className="empty-row">
                      엑셀에서 이름(과 누적 업무강도) 열을 복사해 [붙여넣기] 하세요.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>

      {paste === 'days' && (
        <BulkPaste
          title="고사 일정 붙여넣기"
          help={
            <>
              한 줄에 <b>날짜 시작교시 종료교시</b>를 입력하세요. 탭·쉼표·공백으로 구분합니다. 날짜는 2026-10-12, 10/12, 10월 12일 형식 모두 됩니다.
            </>
          }
          placeholder={'2026-10-12\t1\t3\n2026-10-13\t1\t3\n10/14 1 2'}
          onApply={(lines, mode) => applyPaste('days', lines, mode)}
          onClose={() => setPaste(null)}
        />
      )}
      {paste === 'rooms' && (
        <BulkPaste
          title="고사실 붙여넣기"
          help="한 줄에 하나씩, 또는 엑셀의 한 행을 그대로 복사해 붙여넣으세요."
          placeholder={'1-1\n1-2\n1-3'}
          onApply={(lines, mode) => applyPaste('rooms', lines, mode)}
          onClose={() => setPaste(null)}
        />
      )}
      {paste === 'teachers' && (
        <BulkPaste
          title="감독 교사 붙여넣기"
          help={
            <>
              한 줄에 <b>이름</b> 또는 <b>이름[탭]이전 누적 업무강도</b>를 입력하세요. 엑셀에서 두 열을 복사하면 됩니다.
            </>
          }
          placeholder={'김민준\t350\n이서연\t300\n박도윤'}
          onApply={(lines, mode) => applyPaste('teachers', lines, mode)}
          onClose={() => setPaste(null)}
        />
      )}
      {paste === 'roles' && (
        <BulkPaste
          title="보직 붙여넣기"
          help={
            <>
              한 줄에 <b>보직명[탭]업무강도</b>를 입력하세요.
            </>
          }
          placeholder={'정감독\t100\n부감독\t50\n복도감독\t30'}
          onApply={(lines, mode) => applyPaste('roles', lines, mode)}
          onClose={() => setPaste(null)}
        />
      )}
      {gen && <RoomGenerator onClose={() => setGen(false)} />}
    </div>
  );
}

function RoomGenerator({ onClose }: { onClose: () => void }) {
  const { store, notify } = useApp();
  const [grades, setGrades] = useState('1,2,3');
  const [classes, setClasses] = useState('6');
  const [fmt, setFmt] = useState('{학년}-{반}');
  const [extra, setExtra] = useState('');

  const gradeList = grades.split(/[,\s]+/).filter(Boolean);
  const classList = classes.split(/[,\s]+/).filter(Boolean).map(Number);
  const names: string[] = [];
  gradeList.forEach((g, gi) => {
    const n = classList[gi] ?? classList[classList.length - 1] ?? 0;
    for (let c = 1; c <= n; c++) names.push(fmt.replace('{학년}', g).replace('{반}', String(c)));
  });
  extra
    .split(/[,\n]+/)
    .map((s) => s.trim())
    .filter(Boolean)
    .forEach((s) => names.push(s));

  const apply = (mode: 'append' | 'replace') => {
    store.update((pp) => ({ ...pp, rooms: [...(mode === 'replace' ? [] : pp.rooms), ...names.map((name) => ({ id: uid(), name }))] }), { clean: mode === 'replace' });
    notify(`고사실 ${names.length}개를 만들었습니다.`, 'ok');
    onClose();
  };

  return (
    <Modal
      title="학년·반 고사실 생성"
      onClose={onClose}
      width={520}
      footer={
        <>
          <span className="muted">{names.length}개</span>
          <div className="spacer" />
          <button className="btn" onClick={onClose}>
            취소
          </button>
          <button className="btn danger-outline" disabled={!names.length} onClick={() => apply('replace')}>
            기존 목록 바꾸기
          </button>
          <button className="btn primary" disabled={!names.length} onClick={() => apply('append')}>
            뒤에 추가
          </button>
        </>
      }
    >
      <div className="form-grid">
        <label>
          학년 목록
          <input className="inp" value={grades} onChange={(e) => setGrades(e.target.value)} placeholder="1,2,3" />
        </label>
        <label>
          학년별 반 수 <span className="muted">(하나면 모든 학년 동일)</span>
          <input className="inp" value={classes} onChange={(e) => setClasses(e.target.value)} placeholder="6 또는 7,7,8" />
        </label>
        <label>
          이름 형식
          <input className="inp" value={fmt} onChange={(e) => setFmt(e.target.value)} />
        </label>
        <label>
          추가 고사실 <span className="muted">(쉼표 구분)</span>
          <input className="inp" value={extra} onChange={(e) => setExtra(e.target.value)} placeholder="예: 대기실, 별도고사실" />
        </label>
      </div>
      <div className="preview-chips">
        {names.slice(0, 60).map((n, i) => (
          <span key={i} className="chip">
            {n}
          </span>
        ))}
        {names.length > 60 && <span className="muted">외 {names.length - 60}개</span>}
      </div>
    </Modal>
  );
}
