import { useCallback, useEffect, useRef, useState } from 'react';
import type { Project } from './types';
import { cleanup, emptyProject } from './model';

const STORAGE_KEY = 'proctor-app:project:v1';
const HISTORY_LIMIT = 80;

function load(): Project {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return normalize(JSON.parse(raw));
  } catch {
    /* 저장된 값이 없거나 손상됨 */
  }
  return emptyProject();
}

export function normalize(p: Partial<Project>): Project {
  const base = emptyProject();
  return cleanup({
    ...base,
    ...p,
    version: 1,
    meta: { ...base.meta, ...(p.meta || {}) },
    days: p.days || [],
    rooms: p.rooms || [],
    roles: p.roles || base.roles,
    teachers: (p.teachers || []).map((t) => ({
      ...t,
      prevLoad: t.prevLoad ?? 0,
      target: t.target ?? null,
      forbidden: t.forbidden ?? [],
      availableDays: t.availableDays?.length ? t.availableDays : undefined,
      lockTarget: t.lockTarget || undefined,
    })),
    need: p.need || {},
    cells: p.cells || {},
    extra: p.extra || {},
    exams: p.exams || {},
  });
}

export interface Store {
  project: Project;
  update: (fn: (p: Project) => Project, opts?: { tag?: string; clean?: boolean }) => void;
  replace: (p: Project) => void;
  undo: () => void;
  redo: () => void;
  canUndo: boolean;
  canRedo: boolean;
  savedAt: number | null;
}

export function useStore(): Store {
  const [project, setProject] = useState<Project>(load);
  const past = useRef<Project[]>([]);
  const future = useRef<Project[]>([]);
  const lastTag = useRef<{ tag: string; at: number } | null>(null);
  const [, force] = useState(0);
  const [savedAt, setSavedAt] = useState<number | null>(null);

  useEffect(() => {
    const h = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(project));
        setSavedAt(Date.now());
      } catch {
        /* 저장 공간 부족 등 */
      }
    }, 400);
    return () => clearTimeout(h);
  }, [project]);

  const update = useCallback<Store['update']>((fn, opts) => {
    setProject((prev) => {
      let next = fn(prev);
      if (next === prev) return prev;
      if (opts?.clean) next = cleanup(next);
      const now = Date.now();
      const coalesce = opts?.tag && lastTag.current && lastTag.current.tag === opts.tag && now - lastTag.current.at < 1200;
      if (!coalesce) {
        past.current.push(prev);
        if (past.current.length > HISTORY_LIMIT) past.current.shift();
      }
      lastTag.current = opts?.tag ? { tag: opts.tag, at: now } : null;
      future.current = [];
      return next;
    });
    force((n) => n + 1);
  }, []);

  const replace = useCallback((p: Project) => {
    setProject((prev) => {
      past.current.push(prev);
      future.current = [];
      return normalize(p);
    });
    lastTag.current = null;
    force((n) => n + 1);
  }, []);

  const undo = useCallback(() => {
    setProject((prev) => {
      const p = past.current.pop();
      if (!p) return prev;
      future.current.push(prev);
      return p;
    });
    lastTag.current = null;
    force((n) => n + 1);
  }, []);

  const redo = useCallback(() => {
    setProject((prev) => {
      const p = future.current.pop();
      if (!p) return prev;
      past.current.push(prev);
      return p;
    });
    lastTag.current = null;
    force((n) => n + 1);
  }, []);

  return {
    project,
    update,
    replace,
    undo,
    redo,
    canUndo: past.current.length > 0,
    canRedo: future.current.length > 0,
    savedAt,
  };
}
