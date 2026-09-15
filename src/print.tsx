import { useEffect } from 'react';
import { createPortal } from 'react-dom';
import type { Project, Slot } from './types';
import { docTitle, fmtDay, type Stats } from './model';
import { ChartTable } from './views/ChartView';
import { PersonalCard } from './views/PersonalView';
import { StatsTable } from './views/StatsView';

export type PrintJob =
  | { kind: 'chart'; mode: 'all' | 'daily'; dayIdx?: number; display: 'name' | 'number'; highlightId?: string; hideEmpty: boolean }
  | { kind: 'personal'; teacherIds: string[]; layout: 'page' | 'compact'; hideEmpty: boolean }
  | { kind: 'stats' };

export function PrintHost({ job, project, slots, stats, onDone }: { job: PrintJob; project: Project; slots: Slot[]; stats: Stats; onDone: () => void }) {
  const root = document.getElementById('print-root')!;

  useEffect(() => {
    const landscape = job.kind !== 'personal';
    const style = document.createElement('style');
    style.textContent = `@page { size: A4 ${landscape ? 'landscape' : 'portrait'}; margin: 9mm; }`;
    document.head.appendChild(style);
    const finish = () => {
      window.removeEventListener('afterprint', finish);
      style.remove();
      onDone();
    };
    window.addEventListener('afterprint', finish);
    const h = setTimeout(() => window.print(), 120);
    return () => {
      clearTimeout(h);
      window.removeEventListener('afterprint', finish);
      style.remove();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [job]);

  const title = docTitle(project);
  let content;
  if (job.kind === 'chart') {
    const days = [...new Map(slots.map((s) => [s.dayIdx, s])).values()];
    const dense = project.rooms.length > 16 ? (project.rooms.length > 26 ? '2' : '1') : '0';
    const pages = job.mode === 'all' ? [undefined] : job.dayIdx !== undefined ? [job.dayIdx] : days.map((d) => d.dayIdx);
    content = pages.map((dayIdx, i) => {
      const day = dayIdx === undefined ? undefined : days.find((d) => d.dayIdx === dayIdx);
      return (
        <section key={i} className="print-page print-chart" data-dense={dense}>
          <h1 className="print-title">
            {title} 감독표{day ? ` — ${fmtDay(day.date)}` : ''}
          </h1>
          <ChartTable project={project} slots={slots} dayIdx={dayIdx} display={job.display} highlightId={job.highlightId} hideEmpty={job.hideEmpty} />
        </section>
      );
    });
  } else if (job.kind === 'personal') {
    if (job.layout === 'page') {
      content = job.teacherIds.map((id) => (
        <section key={id} className="print-page print-personal">
          <PersonalCard project={project} slots={slots} teacherId={id} hideEmpty={job.hideEmpty} />
        </section>
      ));
    } else {
      const chunks: string[][] = [];
      for (let i = 0; i < job.teacherIds.length; i += 4) chunks.push(job.teacherIds.slice(i, i + 4));
      content = chunks.map((ids, i) => (
        <section key={i} className="print-page print-compact">
          {ids.map((id) => (
            <PersonalCard key={id} project={project} slots={slots} teacherId={id} hideEmpty compact />
          ))}
        </section>
      ));
    }
  } else {
    content = (
      <section className="print-page print-stats">
        <h1 className="print-title">{title} 감독 현황</h1>
        <StatsTable project={project} slots={slots} stats={stats} />
      </section>
    );
  }

  return createPortal(<div className="print-doc">{content}</div>, root);
}
