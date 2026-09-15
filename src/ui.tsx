import { createContext, useContext, useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import type { Store } from './store';
import type { Slot } from './types';
import type { Issue, Stats } from './model';
import type { PrintJob } from './print';

export type Tab = 'info' | 'need' | 'assign' | 'chart' | 'personal' | 'stats';
export type NotifyKind = 'ok' | 'warn' | 'error' | 'info';

export interface ConfirmOpts {
  title: string;
  message: ReactNode;
  ok?: string;
  cancel?: string;
  danger?: boolean;
}

export interface AppCtx {
  store: Store;
  slots: Slot[];
  stats: Stats;
  issues: Issue[];
  notify: (msg: string, kind?: NotifyKind) => void;
  confirm: (opts: ConfirmOpts) => Promise<boolean>;
  print: (job: PrintJob) => void;
  setTab: (t: Tab) => void;
  saveJson: () => void;
  exportExcel: () => void;
}

export const AppContext = createContext<AppCtx | null>(null);
export function useApp(): AppCtx {
  const c = useContext(AppContext);
  if (!c) throw new Error('AppContext missing');
  return c;
}

export function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

export const safeFileName = (s: string) => s.replace(/[\\/:*?"<>|]/g, '_').trim() || '시감표';

export function Modal({
  title,
  onClose,
  children,
  footer,
  width,
}: {
  title: string;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  width?: number;
}) {
  useEffect(() => {
    const h = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, [onClose]);
  return (
    <div
      className="modal-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="modal" style={{ width }} role="dialog" aria-modal="true" aria-label={title}>
        <div className="modal-head">
          <h3>{title}</h3>
          <button className="icon-btn" onClick={onClose} aria-label="닫기">
            ✕
          </button>
        </div>
        <div className="modal-body">{children}</div>
        {footer && <div className="modal-foot">{footer}</div>}
      </div>
    </div>
  );
}

export function Popover({
  x,
  y,
  onClose,
  children,
  width = 290,
}: {
  x: number;
  y: number;
  onClose: () => void;
  children: ReactNode;
  width?: number;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ left: x, top: y });
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({
      left: Math.max(8, Math.min(x + 6, window.innerWidth - r.width - 8)),
      top: Math.max(8, Math.min(y + 6, window.innerHeight - r.height - 8)),
    });
  }, [x, y]);
  useEffect(() => {
    const down = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    const h = setTimeout(() => window.addEventListener('mousedown', down));
    window.addEventListener('keydown', key);
    return () => {
      clearTimeout(h);
      window.removeEventListener('mousedown', down);
      window.removeEventListener('keydown', key);
    };
  }, [onClose]);
  return (
    <div ref={ref} className="popover" style={{ left: pos.left, top: pos.top, width }}>
      {children}
    </div>
  );
}

export function NumberInput({
  value,
  onChange,
  min,
  max,
  allowEmpty,
  className,
  placeholder,
  title,
  integer,
}: {
  value: number | null;
  onChange: (v: number | null) => void;
  min?: number;
  max?: number;
  allowEmpty?: boolean;
  className?: string;
  placeholder?: string;
  title?: string;
  integer?: boolean;
}) {
  const show = (v: number | null) => (v === null || v === undefined ? '' : String(v));
  const [text, setText] = useState(show(value));
  useEffect(() => {
    setText((t) => (t.trim() !== '' && Number(t) === value ? t : show(value)));
  }, [value]);
  return (
    <input
      className={`inp num ${className ?? ''}`}
      inputMode={integer ? 'numeric' : 'decimal'}
      value={text}
      placeholder={placeholder}
      title={title}
      onChange={(e) => {
        const t = e.target.value;
        setText(t);
        if (t.trim() === '') {
          onChange(allowEmpty ? null : 0);
          return;
        }
        let n = Number(t);
        if (Number.isNaN(n)) return;
        if (integer) n = Math.floor(n);
        if (min !== undefined) n = Math.max(min, n);
        if (max !== undefined) n = Math.min(max, n);
        onChange(n);
      }}
      onBlur={() => setText(show(value))}
    />
  );
}

export function BulkPaste({
  title,
  help,
  placeholder,
  onApply,
  onClose,
  allowReplace = true,
}: {
  title: string;
  help: ReactNode;
  placeholder?: string;
  onApply: (lines: string[], mode: 'append' | 'replace') => void;
  onClose: () => void;
  allowReplace?: boolean;
}) {
  const [text, setText] = useState('');
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.replace(/\s+$/, ''))
    .filter((l) => l.trim());
  return (
    <Modal
      title={title}
      onClose={onClose}
      width={560}
      footer={
        <>
          <span className="muted">{lines.length}줄 인식</span>
          <div className="spacer" />
          <button className="btn" onClick={onClose}>
            취소
          </button>
          {allowReplace && (
            <button
              className="btn danger-outline"
              disabled={!lines.length}
              onClick={() => {
                onApply(lines, 'replace');
                onClose();
              }}
            >
              기존 목록 바꾸기
            </button>
          )}
          <button
            className="btn primary"
            disabled={!lines.length}
            onClick={() => {
              onApply(lines, 'append');
              onClose();
            }}
          >
            뒤에 추가
          </button>
        </>
      }
    >
      <div className="help">{help}</div>
      <textarea className="inp paste" autoFocus rows={12} value={text} placeholder={placeholder} onChange={(e) => setText(e.target.value)} />
    </Modal>
  );
}

export function IssueList({ issues, max = 40, empty }: { issues: Issue[]; max?: number; empty?: ReactNode }) {
  if (!issues.length) return <div className="issues-empty">✓ {empty ?? '확인된 문제가 없습니다.'}</div>;
  const sorted = [...issues].sort((a, b) => (a.level === b.level ? 0 : a.level === 'error' ? -1 : 1));
  return (
    <ul className="issues">
      {sorted.slice(0, max).map((i, k) => (
        <li key={k} className={`issue ${i.level}`}>
          <span className="dot" />
          {i.msg}
        </li>
      ))}
      {sorted.length > max && <li className="issue more">외 {sorted.length - max}건</li>}
    </ul>
  );
}

export function Seg<T extends string>({ value, options, onChange }: { value: T; options: { v: T; label: string }[]; onChange: (v: T) => void }) {
  return (
    <div className="seg" role="group">
      {options.map((o) => (
        <button key={o.v} className={value === o.v ? 'on' : ''} onClick={() => onChange(o.v)} aria-pressed={value === o.v}>
          {o.label}
        </button>
      ))}
    </div>
  );
}
