import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { normalize, useStore } from './store';
import { computeStats, docTitle, emptyProject, getSlots, remapPreserve, validate, type IssueArea } from './model';
import { demoProject } from './demo';
import { AppContext, Modal, download, safeFileName, type AppCtx, type ConfirmOpts, type NotifyKind, type Tab } from './ui';
import { PrintHost, type PrintJob } from './print';
import { downloadTemplate, exportWorkbook, importFromWorkbook, isProjectWorkbook, openWorkbook } from './excel';
import { applyTimetable, parseTimetable, type ParsedTimetable } from './timetable';
import InfoView from './views/InfoView';
import NeedView from './views/NeedView';
import AssignView from './views/AssignView';
import ChartView from './views/ChartView';
import PersonalView from './views/PersonalView';
import StatsView from './views/StatsView';

const TABS: { id: Tab; label: string; sub: string; areas: IssueArea[] }[] = [
  { id: 'info', label: '기본 정보', sub: '일정·고사실·교사·보직', areas: ['info'] },
  { id: 'need', label: '필요 감독 수', sub: '고사실별 인원', areas: ['need'] },
  { id: 'assign', label: '감독 배정', sub: '불가·고정·자동 배정', areas: ['assign', 'result'] },
  { id: 'chart', label: '감독표', sub: '전체·일별', areas: [] },
  { id: 'personal', label: '개인 시간표', sub: '교사별 인쇄', areas: [] },
  { id: 'stats', label: '감독 현황', sub: '업무강도 통계', areas: [] },
];

interface Toast {
  id: number;
  msg: string;
  kind: NotifyKind;
}

export default function App() {
  const store = useStore();
  const { project } = store;
  const slots = useMemo(() => getSlots(project), [project]);
  const stats = useMemo(() => computeStats(project, slots), [project, slots]);
  const issues = useMemo(() => validate(project, slots, stats), [project, slots, stats]);

  const [tab, setTabState] = useState<Tab>(() => {
    const saved = localStorage.getItem('proctor-app:tab') as Tab | null;
    return saved && TABS.some((t) => t.id === saved) ? saved : 'info';
  });
  const setTab = useCallback((t: Tab) => {
    setTabState(t);
    localStorage.setItem('proctor-app:tab', t);
    window.scrollTo({ top: 0 });
  }, []);

  const [toasts, setToasts] = useState<Toast[]>([]);
  const notify = useCallback((msg: string, kind: NotifyKind = 'info') => {
    const id = Date.now() + Math.random();
    setToasts((ts) => [...ts, { id, msg, kind }]);
    setTimeout(() => setToasts((ts) => ts.filter((t) => t.id !== id)), kind === 'error' ? 7000 : 3800);
  }, []);

  const [confirmState, setConfirmState] = useState<{ opts: ConfirmOpts; resolve: (v: boolean) => void } | null>(null);
  const confirm = useCallback(
    (opts: ConfirmOpts) => new Promise<boolean>((resolve) => setConfirmState({ opts, resolve })),
    [],
  );
  const closeConfirm = (v: boolean) => {
    confirmState?.resolve(v);
    setConfirmState(null);
  };

  const [printJob, setPrintJob] = useState<PrintJob | null>(null);
  const [help, setHelp] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  const fileBase = safeFileName(docTitle(project) || '시감표');
  const saveJson = useCallback(() => {
    download(new Blob([JSON.stringify(project)], { type: 'application/json' }), `${fileBase}.json`);
    notify('프로젝트 파일(JSON)을 저장했습니다.', 'ok');
  }, [project, fileBase, notify]);

  const exportExcel = useCallback(async () => {
    notify('엑셀 파일을 만드는 중입니다…');
    try {
      await exportWorkbook(project, slots, stats);
      notify('엑셀 파일을 저장했습니다.', 'ok');
    } catch (e) {
      notify(`엑셀 내보내기 실패: ${(e as Error).message}`, 'error');
    }
  }, [project, slots, stats, notify]);

  const openFile = useCallback(() => fileRef.current?.click(), []);

  const getTemplate = useCallback(async () => {
    notify('엑셀 입력 양식을 만드는 중입니다…');
    try {
      await downloadTemplate(project, slots);
      notify('입력 양식을 저장했습니다. 엑셀에서 채운 뒤 [작성한 엑셀 올리기]로 올리세요.', 'ok');
    } catch (e) {
      notify(`양식 만들기 실패: ${(e as Error).message}`, 'error');
    }
  }, [project, slots, notify]);

  /** 시험 시간표(PDF·엑셀)에서 읽은 날짜·교시를 현재 작업에 적용 */
  const applyParsedTimetable = async (tt: ParsedTimetable, fileName: string) => {
    const { project: next, notes } = applyTimetable(project, tt);
    const ok = await confirm({
      title: '시험 시간표 불러오기',
      message: (
        <>
          <p>“{fileName}”에서 시험 일정을 읽었습니다. 고사 날짜와 교시를 아래와 같이 설정합니다. (실행 취소로 되돌릴 수 있습니다)</p>
          <ul className="notes">
            {notes.map((n, i) => (
              <li key={i}>{n}</li>
            ))}
          </ul>
        </>
      ),
      ok: '일정 적용',
    });
    if (!ok) return;
    store.update(() => next);
    setTab('info');
    notify('시험 날짜와 교시를 설정했습니다.', 'ok');
  };

  const onFile = async (file: File) => {
    const name = file.name.toLowerCase();
    try {
      if (name.endsWith('.pdf')) {
        const { parseTimetablePdf } = await import('./timetablePdf');
        const tt = await parseTimetablePdf(file);
        if (!tt) throw new Error('PDF에서 시험 시간표(교시와 날짜 표)를 찾지 못했습니다. 학교 시험 시간표 PDF인지 확인해 주세요.');
        await applyParsedTimetable(tt, file.name);
        return;
      }
      if (name.endsWith('.json')) {
        const data = JSON.parse(await file.text());
        if (!data || !Array.isArray(data.teachers)) throw new Error('시감표 프로젝트 파일이 아닙니다.');
        const ok = await confirm({ title: '프로젝트 불러오기', message: `현재 작업을 "${file.name}" 내용으로 바꿉니다. (실행 취소로 되돌릴 수 있습니다)`, ok: '불러오기' });
        if (ok) {
          store.replace(normalize(data));
          notify('프로젝트를 불러왔습니다.', 'ok');
        }
      } else if (/\.(xlsx|xlsm)$/.test(name)) {
        const wb = await openWorkbook(file);
        if (!isProjectWorkbook(wb)) {
          const tt = parseTimetable(wb);
          if (!tt) throw new Error("'시험기본정보' 시트나 시험 시간표(교시와 날짜 표)를 찾지 못했습니다. proctor 엑셀, 이 앱의 양식, 시험 시간표 중 하나를 선택하세요.");
          await applyParsedTimetable(tt, file.name);
          return;
        }
        const { project: imported, notes } = importFromWorkbook(wb);
        const merged = remapPreserve(project, imported);
        const keptNeed = Object.keys(merged.need).length - Object.keys(imported.need).length;
        const keptCells = Object.keys(merged.cells).length - Object.keys(imported.cells).length;
        if (keptNeed > 0) notes.push(`파일에 없는 필요 감독 수 ${keptNeed}칸은 지금 작업에서 그대로 가져옵니다.`);
        if (keptCells > 0) notes.push(`파일에 없는 감독 불가·고정 표시 ${keptCells}칸은 지금 작업에서 그대로 가져옵니다.`);
        const ok = await confirm({
          title: '엑셀 파일 불러오기',
          message: (
            <>
              <p>현재 작업을 “{file.name}” 내용으로 바꿉니다. (실행 취소로 되돌릴 수 있습니다)</p>
              <ul className="notes">
                {notes.map((n, i) => (
                  <li key={i}>{n}</li>
                ))}
              </ul>
            </>
          ),
          ok: '불러오기',
        });
        if (ok) {
          store.replace(merged);
          setTab(Object.keys(merged.need).length ? 'assign' : 'need');
          notify('엑셀 파일을 불러왔습니다.', 'ok');
        }
      } else notify('JSON, 엑셀(xlsx, xlsm), 시험 시간표 PDF 중 하나를 선택하세요.', 'warn');
    } catch (e) {
      notify(`불러오기 실패: ${(e as Error).message}`, 'error');
    }
  };

  useEffect(() => {
    const h = (e: KeyboardEvent) => {
      const el = e.target as HTMLElement;
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT')) return;
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.code === 'KeyZ' && !e.shiftKey) {
        e.preventDefault();
        store.undo();
      } else if (e.code === 'KeyY' || (e.code === 'KeyZ' && e.shiftKey)) {
        e.preventDefault();
        store.redo();
      } else if (e.code === 'KeyS') {
        e.preventDefault();
        saveJson();
      }
    };
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, [store, saveJson]);

  const ctx: AppCtx = { store, slots, stats, issues, notify, confirm, print: setPrintJob, setTab, saveJson, exportExcel, openFile, downloadTemplate: getTemplate };

  const badge = (areas: IssueArea[]) => issues.filter((i) => i.level === 'error' && areas.includes(i.area)).length;

  let view: ReactNode;
  switch (tab) {
    case 'info':
      view = <InfoView />;
      break;
    case 'need':
      view = <NeedView />;
      break;
    case 'assign':
      view = <AssignView />;
      break;
    case 'chart':
      view = <ChartView />;
      break;
    case 'personal':
      view = <PersonalView />;
      break;
    case 'stats':
      view = <StatsView />;
      break;
  }

  return (
    <AppContext.Provider value={ctx}>
      <div
        className={`app ${dragOver ? 'drag-over' : ''}`}
        onDragOver={(e) => {
          if (e.dataTransfer.types.includes('Files')) {
            e.preventDefault();
            setDragOver(true);
          }
        }}
        onDragLeave={(e) => {
          if (e.currentTarget === e.target) setDragOver(false);
        }}
        onDrop={(e) => {
          e.preventDefault();
          setDragOver(false);
          const f = e.dataTransfer.files[0];
          if (f) onFile(f);
        }}
      >
        <header className="topbar">
          <div className="brand">
            <span className="logo" aria-hidden>
              📋
            </span>
            <div>
              <div className="brand-name">정기고사 시감 배정 생성 시스템</div>
              <div className="brand-sub">
                {docTitle(project)
                  ? `${docTitle(project)} · 고등학교 정기고사 감독(시감)을 업무 강도에 따라 공정하게 배정합니다.`
                  : '고등학교 정기고사 감독(시감)을 업무 강도에 따라 공정하게 배정합니다.'}
              </div>
            </div>
          </div>
          <div className="top-actions">
            <button className="btn ghost sm" onClick={store.undo} disabled={!store.canUndo} title="실행 취소 (Ctrl+Z)">
              ↶ 취소
            </button>
            <button className="btn ghost sm" onClick={store.redo} disabled={!store.canRedo} title="다시 실행 (Ctrl+Y)">
              ↷ 다시
            </button>
            <span className="sep" />
            <button
              className="btn ghost sm"
              onClick={async () => {
                if (await confirm({ title: '새로 만들기', message: '모든 입력 내용을 비우고 새로 시작합니다. 필요하면 먼저 [저장]하세요.', ok: '새로 만들기', danger: true }))
                  store.replace(emptyProject());
              }}
            >
              새로 만들기
            </button>
            <button
              className="btn ghost sm"
              onClick={async () => {
                if (await confirm({ title: '예시 데이터', message: '4일·18개 고사실·52명 예시 데이터로 바꿉니다. 현재 작업은 실행 취소로 되돌릴 수 있습니다.', ok: '예시 불러오기' })) {
                  store.replace(demoProject());
                  setTab('info');
                }
              }}
            >
              예시 데이터
            </button>
            <button className="btn ghost sm" onClick={() => fileRef.current?.click()} title="JSON 프로젝트 또는 proctor 엑셀(xlsm/xlsx)">
              열기
            </button>
            <button className="btn ghost sm" onClick={saveJson} title="프로젝트 파일(JSON) 저장 (Ctrl+S)">
              저장
            </button>
            <button className="btn accent sm" onClick={exportExcel}>
              엑셀 내보내기
            </button>
            <button className="btn ghost sm" onClick={() => setHelp(true)}>
              도움말
            </button>
            <input
              ref={fileRef}
              type="file"
              accept=".json,.xlsx,.xlsm,.pdf"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) onFile(f);
                e.target.value = '';
              }}
            />
          </div>
        </header>

        <nav className="tabs" aria-label="작업 단계">
          {TABS.map((t, i) => {
            const n = badge(t.areas);
            return (
              <button key={t.id} className={`tab ${tab === t.id ? 'active' : ''}`} onClick={() => setTab(t.id)}>
                <span className="step">{i + 1}</span>
                <span className="tab-text">
                  <span className="tab-label">
                    {t.label}
                    {n > 0 && <span className="badge">{n}</span>}
                  </span>
                  <span className="tab-sub">{t.sub}</span>
                </span>
              </button>
            );
          })}
          <div className="spacer" />
          <span className="saved">{store.savedAt ? '브라우저에 자동 저장됨' : ''}</span>
        </nav>

        <main className="main">{view}</main>

        <div className="toasts" aria-live="polite">
          {toasts.map((t) => (
            <div key={t.id} className={`toast ${t.kind}`}>
              {t.msg}
            </div>
          ))}
        </div>

        {confirmState && (
          <Modal
            title={confirmState.opts.title}
            onClose={() => closeConfirm(false)}
            width={460}
            footer={
              <>
                <div className="spacer" />
                <button className="btn" onClick={() => closeConfirm(false)}>
                  {confirmState.opts.cancel ?? '취소'}
                </button>
                <button className={`btn ${confirmState.opts.danger ? 'danger' : 'primary'}`} autoFocus onClick={() => closeConfirm(true)}>
                  {confirmState.opts.ok ?? '확인'}
                </button>
              </>
            }
          >
            <div className="confirm-msg">{confirmState.opts.message}</div>
          </Modal>
        )}

        {help && <HelpModal onClose={() => setHelp(false)} />}
      </div>
      {printJob && <PrintHost job={printJob} project={project} slots={slots} stats={stats} onDone={() => setPrintJob(null)} />}
    </AppContext.Provider>
  );
}

function HelpModal({ onClose }: { onClose: () => void }) {
  return (
    <Modal title="사용 방법" onClose={onClose} width={680}>
      <div className="help-doc">
        <ol>
          <li>
            <b>기본 정보</b> — 고사 날짜와 교시, 고사실, 감독 교사(이전 시험까지의 누적 업무강도), 보직과 업무강도를 입력합니다. 엑셀에서 복사해
            [붙여넣기]로 한꺼번에 넣을 수 있습니다.
          </li>
          <li>
            <b>필요 감독 수</b> — 교시·보직별로 각 고사실에 필요한 감독 수를 입력합니다. [빠른 채우기]로 “모든 교시·모든 고사실 정감독 1명”을 한 번에 넣을
            수 있고, 엑셀 표를 셀에 붙여넣어도 됩니다.
          </li>
          <li>
            <b>감독 배정</b> — 감독할 수 없는 시간은 <kbd>X</kbd>, 반드시 감독할 시간은 <kbd>1</kbd>(고정)로 표시하고 교사별 “못 들어가는 고사실”을
            지정합니다. [배정할 시간 자동 채우기] → [감독 배정 시작]을 누르면 자동으로 배정됩니다.
          </li>
          <li>
            <b>감독표 · 개인 시간표</b> — 전체/일별 감독표와 교사별 개인 시간표를 확인하고 인쇄(PDF 저장 포함)하거나 엑셀로 내보냅니다.
          </li>
          <li>
            <b>감독 현황</b> — 교사별 감독 시간, 보직 수, 업무강도를 확인합니다. [다음 고사 준비]로 누적 업무강도를 다음 시험에 이월합니다.
          </li>
        </ol>
        <h4>시험 시간표 불러오기 · 시험 당일 구글 시트</h4>
        <p>
          [기본 정보]의 <b>[시험 시간표 불러오기]</b>에 학교 시험 시간표 PDF(또는 구글 시트를 xlsx로 받은 파일)를 올리면 시험 날짜와 교시가 자동으로
          설정되고, 과목·출제 교사가 저장됩니다. 배정이 끝나면 [감독표] 화면의 <b>시험 당일 공유용 시감표</b>에서 날짜를 골라{' '}
          <b>[구글 시트 만들기]</b>로 학교 양식(시감표 · 교사별 · 감독 누계 · 시험 시간표)의 구글 시트를 만들 수 있습니다. 처음 한 번 [구글 연결
          설정]이 필요하며, 연결 전에는 같은 양식의 엑셀을 받아 드라이브에 올리면 됩니다.
        </p>
        <h4>엑셀로 한 번에 입력하기</h4>
        <p>
          화면에서 하나씩 입력하는 대신, [기본 정보] 화면의 <b>[엑셀 양식 내려받기]</b>로 양식을 받아 채운 뒤 <b>[작성한 엑셀 올리기]</b>로 올리면 한
          번에 입력됩니다. 기본 정보를 먼저 올리고 양식을 다시 받으면, 교시·보직 줄과 교사 명단이 채워진 [배정감독수정보] · [감독배정] 시트가 들어 있어
          필요 감독 수와 감독 불가(x) · 고정(1) 표시까지 엑셀에서 작성할 수 있습니다. 양식을 다시 올려도 파일에 없는 내용은 이름과 날짜를 기준으로
          그대로 유지됩니다.
        </p>
        <h4>자동 배정 방식</h4>
        <p>
          남은 필요 인원 비율(P값)이 가장 높은 교시부터 채우고, 막히면 교사 간 교환 경로를 찾아 해결합니다. 이후 날짜별 감독 수를 고르게 하고, 같은 날
          연속 감독을 줄이며, 누적 업무강도가 낮은 교사에게 무거운 보직을 먼저 배정합니다. 고사실은 무작위로 배정하되 못 들어가는 고사실은 교환으로
          피하고, 마지막으로 교사 간 보직을 맞바꿔 업무강도를 고르게 합니다. 여러 번 시도해 가장 좋은 결과를 고릅니다.
        </p>
        <h4>감독 배정 표 단축키</h4>
        <ul className="keys">
          <li>
            드래그 / <kbd>Shift</kbd>+클릭: 범위 선택, <kbd>Ctrl</kbd>+클릭: 여러 칸 선택, 날짜·교시·이름 머리글 클릭: 줄 전체 선택
          </li>
          <li>
            <kbd>X</kbd> 감독 불가 · <kbd>1</kbd> 또는 <kbd>F</kbd> 고정 배정 · <kbd>Delete</kbd> 지우기 · <kbd>Enter</kbd>/더블클릭 칸 편집
          </li>
          <li>
            <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>C</kbd>: 선택한 두 칸 교환 (두 교사 × 두 시간 4칸을 선택하면 두 교사의 감독 시간을 맞바꿈)
          </li>
          <li>
            <kbd>Ctrl</kbd>+<kbd>Z</kbd> 실행 취소 · <kbd>Ctrl</kbd>+<kbd>Y</kbd> 다시 실행 · <kbd>Ctrl</kbd>+<kbd>S</kbd> 프로젝트 저장
          </li>
        </ul>
        <h4>저장</h4>
        <p>
          작업 내용은 이 브라우저에 자동 저장됩니다. 다른 컴퓨터로 옮기거나 보관하려면 [저장]으로 JSON 파일을 받아 두세요. [열기]로 기존 proctor
          엑셀 파일(xlsm)이나 이 앱에서 내보낸 엑셀 파일도 불러올 수 있습니다.
        </p>
        <p className="muted">원본: proctor_ver2026 엑셀 프로그램 (namgungyeon.tistory.com/54)</p>
      </div>
    </Modal>
  );
}
