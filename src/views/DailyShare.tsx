import { useMemo, useState } from 'react';
import { Modal, download, safeFileName, useApp } from '../ui';
import { fmtDay } from '../model';
import { buildDayBook } from '../daySheet';
import { sheetBookToXlsx } from '../excel';
import {
  appsScriptCode,
  createGoogleSheet,
  isWebAppUrl,
  loadSheetSettings,
  randomToken,
  saveSheetSettings,
  testConnection,
  type SheetSettings,
} from '../appsScript';

const LINKS_KEY = 'proctor-app:sheet-links';

const todayIso = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

async function copyText(text: string) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}

/** 시험 당일 아침 교사들에게 공유할 날짜별 시감표(구글 시트 / 엑셀) */
export default function DailyShareCard() {
  const { store, slots, stats, notify } = useApp();
  const p = store.project;
  const days = useMemo(() => [...new Map(slots.map((s) => [s.dayIdx, s])).values()], [slots]);
  const [picked, setPicked] = useState<number | null>(null);
  const [settings, setSettings] = useState<SheetSettings>(loadSheetSettings);
  const [busy, setBusy] = useState(false);
  const [setup, setSetup] = useState(false);
  const [links, setLinks] = useState<Record<string, string>>(() => {
    try {
      return JSON.parse(localStorage.getItem(LINKS_KEY) || '{}');
    } catch {
      return {};
    }
  });

  if (!days.length) return null;
  // 오늘이 시험일이면 오늘, 아니면 다가오는 첫 시험일을 기본으로
  const today = todayIso();
  const defaultDay = days.find((d) => d.date === today) ?? days.find((d) => d.date > today) ?? days[0];
  const day = days.find((d) => d.dayIdx === picked) ?? defaultDay;
  const linkKey = `${p.meta.title}|${day.date}`;
  const connected = isWebAppUrl(settings.url);
  const hasExams = Object.keys(p.exams ?? {}).length > 0;

  const makeSheet = async () => {
    setBusy(true);
    try {
      const book = buildDayBook(p, slots, day.dayIdx);
      const res = await createGoogleSheet(settings, book);
      const next = { ...links, [linkKey]: res.url };
      setLinks(next);
      try {
        localStorage.setItem(LINKS_KEY, JSON.stringify(next));
      } catch {
        /* 저장 불가 */
      }
      const copied = await copyText(res.url);
      notify(`“${book.title}” 구글 시트를 만들었습니다.${copied ? ' 공유 링크를 복사했습니다.' : ''}`, 'ok');
    } catch (e) {
      notify(`구글 시트 만들기 실패: ${(e as Error).message}`, 'error');
    } finally {
      setBusy(false);
    }
  };

  const downloadXlsx = async () => {
    try {
      const book = buildDayBook(p, slots, day.dayIdx);
      download(await sheetBookToXlsx(book), `${safeFileName(book.title)}.xlsx`);
      notify('엑셀 파일을 저장했습니다. 구글 드라이브에 올려 “Google 스프레드시트로 열기”를 하면 됩니다.', 'ok');
    } catch (e) {
      notify(`엑셀 만들기 실패: ${(e as Error).message}`, 'error');
    }
  };

  return (
    <section className="card share-card">
      <div className="share-main">
        <h3>시험 당일 공유용 시감표</h3>
        <p>
          날짜를 고르면 그날의 시감표(학급별 · 교사별 · 감독 누계{hasExams ? ' · 시험 시간표' : ''})를 학교 양식 그대로 구글 시트로 만듭니다.
          {!hasExams && ' 기본 정보에서 시험 시간표를 불러오면 교시별 과목과 대표 교사도 함께 들어갑니다.'}
        </p>
        <div className="seg share-days">
          {days.map((d) => (
            <button key={d.dayIdx} className={d.dayIdx === day.dayIdx ? 'on' : ''} onClick={() => setPicked(d.dayIdx)}>
              {d.dayIdx + 1}일차 {fmtDay(d.date)}
              {d.date === today ? ' · 오늘' : ''}
            </button>
          ))}
        </div>
        {!stats.hasResults && <p className="warn-text">아직 감독 배정 결과가 없습니다. 감독 배정을 먼저 해 주세요.</p>}
        {links[linkKey] && (
          <div className="share-link">
            만든 시트:{' '}
            <a href={links[linkKey]} target="_blank" rel="noreferrer">
              {day.dayIdx + 1}일차 시감표 열기 ↗
            </a>
            <button
              className="btn sm ghost"
              onClick={async () => notify((await copyText(links[linkKey])) ? '링크를 복사했습니다.' : '복사하지 못했습니다. 시트를 열어 주소를 복사하세요.', 'ok')}
            >
              링크 복사
            </button>
          </div>
        )}
      </div>
      <div className="share-actions">
        <button className="btn primary" disabled={!connected || busy || !stats.hasResults} onClick={makeSheet} title={connected ? '' : '먼저 구글 연결 설정을 해 주세요'}>
          {busy ? '만드는 중…' : '구글 시트 만들기'}
        </button>
        <button className="btn" onClick={downloadXlsx} disabled={!stats.hasResults}>
          같은 양식 엑셀로 받기
        </button>
        <button className="btn sm ghost" onClick={() => setSetup(true)}>
          {connected ? '⚙ 구글 연결 설정' : '⚙ 구글 연결 설정 (처음 1회)'}
        </button>
      </div>
      {setup && (
        <SheetSetupModal
          settings={settings}
          onSave={(s) => {
            setSettings(s);
            saveSheetSettings(s);
          }}
          onClose={() => setSetup(false)}
        />
      )}
    </section>
  );
}

function SheetSetupModal({ settings, onSave, onClose }: { settings: SheetSettings; onSave: (s: SheetSettings) => void; onClose: () => void }) {
  const { notify } = useApp();
  const [s, setS] = useState<SheetSettings>(() => ({ ...settings, token: settings.token || randomToken() }));
  const [testing, setTesting] = useState(false);
  const code = appsScriptCode(s.token);

  const test = async () => {
    setTesting(true);
    try {
      await testConnection(s.url);
      notify('연결되었습니다. [저장]을 누르세요.', 'ok');
    } catch (e) {
      notify(`연결 실패: ${(e as Error).message}`, 'error');
    } finally {
      setTesting(false);
    }
  };

  return (
    <Modal
      title="구글 시트 연결 설정"
      onClose={onClose}
      width={780}
      footer={
        <>
          <button className="btn" onClick={test} disabled={!isWebAppUrl(s.url) || testing}>
            {testing ? '확인 중…' : '연결 테스트'}
          </button>
          <div className="spacer" />
          <button className="btn" onClick={onClose}>
            닫기
          </button>
          <button
            className="btn primary"
            onClick={() => {
              onSave(s);
              notify('설정을 저장했습니다.', 'ok');
              onClose();
            }}
          >
            저장
          </button>
        </>
      }
    >
      <p className="muted">처음 한 번만 하면 됩니다. 시트는 이 스크립트를 배포한 구글 계정의 드라이브에 만들어집니다.</p>
      <ol className="setup-steps">
        <li>
          <a href="https://script.google.com/home/projects/create" target="_blank" rel="noreferrer">
            script.google.com 새 프로젝트
          </a>
          를 엽니다. (시감표를 보관할 구글 계정으로 로그인)
        </li>
        <li>
          기본 코드를 모두 지우고 아래 <b>코드를 붙여넣은 뒤 저장</b>합니다. 코드에는 이 앱 전용 연결 토큰이 들어 있습니다.
        </li>
        <li>
          오른쪽 위 <b>배포 → 새 배포</b> → 유형 <b>웹 앱</b> → 실행 사용자 <b>나</b>, 액세스 권한 <b>모든 사용자</b> → <b>배포</b>. 권한 승인 창에서
          “확인되지 않은 앱”이 나오면 <b>고급 → (프로젝트 이름)(으)로 이동</b>을 누릅니다.
        </li>
        <li>
          나온 <b>웹 앱 URL</b>(…/exec)을 아래에 붙여넣고 <b>[연결 테스트] → [저장]</b>.
        </li>
      </ol>

      <div className="code-box">
        <div className="code-head">
          <span>Code.gs</span>
          <button className="btn sm" onClick={async () => notify((await copyText(code)) ? '코드를 복사했습니다.' : '복사하지 못했습니다. 코드 칸을 눌러 전체 선택 후 복사하세요.', 'ok')}>
            코드 복사
          </button>
        </div>
        <textarea readOnly className="code" value={code} rows={9} onFocus={(e) => e.currentTarget.select()} />
      </div>

      <div className="form-grid">
        <label>
          웹 앱 URL
          <input className="inp" value={s.url} placeholder="https://script.google.com/macros/s/…/exec" onChange={(e) => setS({ ...s, url: e.target.value })} />
        </label>
        <label>
          저장할 드라이브 폴더 <span className="muted">(선택 · 폴더 주소나 ID)</span>
          <input className="inp" value={s.folderId} placeholder="비워 두면 내 드라이브 첫 화면" onChange={(e) => setS({ ...s, folderId: e.target.value })} />
        </label>
        <label>
          연결 토큰 <span className="muted">(코드 안의 값과 같아야 함)</span>
          <span className="token-row">
            <input className="inp" value={s.token} onChange={(e) => setS({ ...s, token: e.target.value })} />
            <button className="btn sm" onClick={() => setS({ ...s, token: randomToken() })}>
              새로 만들기
            </button>
          </span>
        </label>
        <label className="pop-check share-check">
          <input type="checkbox" checked={s.shareLink} onChange={(e) => setS({ ...s, shareLink: e.target.checked })} />
          만든 시트를 “링크가 있는 모든 사용자 — 보기”로 공유
        </label>
      </div>
      <p className="hint">
        토큰을 바꾸면 코드를 다시 복사해 붙여넣고 <b>배포 → 배포 관리 → 새 버전</b>으로 다시 배포해야 합니다. 웹 앱 URL과 토큰은 이 브라우저에만
        저장됩니다. 링크 공유를 켜면 주소를 아는 사람은 누구나 볼 수 있으니 교사 단체방에만 공유하세요.
      </p>
    </Modal>
  );
}
