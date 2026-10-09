// Modal dialogs for the complex element types (tables, charts, shapes, media).
//
// This module is deliberately model-agnostic: it only builds DOM and reports the user's intent
// back through callbacks, so `main.ts` stays the single owner of the store/actions.
import { h } from './util';
import { SHAPES, CHART_TYPES } from '@fefeding/ppt-parser';

let activeModal: HTMLElement | null = null;

export function hasOpenModal(): boolean {
  return !!activeModal;
}

/** Close the current dialog, if any. Bound to Escape and to the backdrop. */
export function closeModal(): void {
  activeModal?.remove();
  activeModal = null;
}

// ---------- tiny helpers ----------
function input(type: string, value: string, cls = 'dlg-in'): HTMLInputElement {
  const el = document.createElement('input');
  el.type = type;
  el.value = value;
  el.className = cls;
  return el;
}
function select(options: { value: string; name: string }[], value: string, cls = 'dlg-in'): HTMLSelectElement {
  const el = document.createElement('select');
  el.className = cls;
  for (const o of options) {
    const opt = document.createElement('option');
    opt.value = o.value;
    opt.textContent = o.name;
    if (o.value === value) opt.selected = true;
    el.append(opt);
  }
  return el;
}
function intOf(v: string, min: number, max: number, def: number): number {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : def;
}
function btn(label: string, onclick: () => void, cls = 'btn'): HTMLElement {
  return h('button', { class: cls, onclick }, [label]);
}
function row(nodes: HTMLElement[], cls = 'dlg-row'): HTMLElement {
  return h('div', { class: cls }, nodes);
}
function labelled(label: string, node: HTMLElement): HTMLElement {
  return h('label', { class: 'dlg-line' }, [h('span', {}, [label]), node]);
}
function checkBox(label: string, value: boolean, onChange: (v: boolean) => void): HTMLElement {
  const box = input('checkbox', '', 'dlg-check');
  box.checked = value;
  box.addEventListener('change', () => onChange(box.checked));
  return h('label', { class: 'dlg-check-row' }, [box, h('span', {}, [label])]);
}

// ---------- dialog shell ----------
function openDialog(title: string, build: (close: () => void) => HTMLElement[]): void {
  closeModal();
  const inner = h('div', { class: 'dlg' }, [
    h('div', { class: 'dlg-head' }, [h('span', {}, [title]), h('span', { class: 'dlg-x', onclick: closeModal }, ['✕'])])
  ]);
  const modal = h('div', { class: 'modal' }, [inner]);
  modal.addEventListener('pointerdown', (e) => {
    if (e.target === modal) closeModal();
  });
  document.body.append(modal);
  activeModal = modal;
  const body = h('div', { class: 'dlg-body' });
  build(closeModal).forEach((n) => body.append(n));
  inner.append(body);
}

// ---------- table: insert ----------
export function openTableSizeDialog(api: { create: (rows: number, cols: number) => void }): void {
  openDialog('Insert table', (close) => {
    const rowsIn = input('number', '3');
    rowsIn.min = '1';
    rowsIn.max = '30';
    const colsIn = input('number', '3');
    colsIn.min = '1';
    colsIn.max = '20';
    return [
      labelled('Rows', rowsIn),
      labelled('Columns', colsIn),
      row([
        btn('Cancel', close),
        btn('Insert', () => {
          api.create(intOf(rowsIn.value, 1, 30, 3), intOf(colsIn.value, 1, 20, 3));
          close();
        }, 'btn primary')
      ], 'dlg-actions')
    ];
  });
}

// ---------- table: edit ----------
export function openTableDialog(
  el: any,
  api: {
    update: (patch: any) => void;
    resize: (rows: number, cols: number) => void;
    reopen: () => void;
  }
): void {
  openDialog('Edit table', (close) => {
    const rows: any[] = Array.isArray(el.rows) ? el.rows : [];
    const cols = Math.max(1, ...rows.map((r: any) => (r.cells || []).length));
    const draft: string[][] = rows.map((r: any) =>
      Array.from({ length: cols }, (_, j) => String((((r.cells || [])[j] || {}) as any).text ?? ''))
    );

    const rowsIn = input('number', String(rows.length || 1), 'dlg-cell');
    const colsIn = input('number', String(cols), 'dlg-cell');

    const grid = h('div', { class: 'dlg-table' });
    draft.forEach((r, i) =>
      r.forEach((v, j) => {
        const cell = input('text', v, 'dlg-cell');
        cell.addEventListener('input', () => {
          draft[i][j] = cell.value;
        });
        grid.append(cell);
      })
    );

    return [
      row([labelled('Rows', rowsIn), labelled('Cols', colsIn), btn('Resize', () => {
        api.resize(intOf(rowsIn.value, 1, 30, 1), intOf(colsIn.value, 1, 20, 1));
        // The action rebuilds this table's rows, so reopen against the fresh model
        close();
        api.reopen();
      })]),
      grid,
      row([
        btn('Cancel', close),
        btn('Save', () => {
          const next = draft.map((r, i) => ({
            height: (rows[i] || {}).height,
            cells: r.map((text, j) => ({ ...((rows[i] || {}).cells || [])[j], text }))
          }));
          api.update({ rows: next });
          close();
        }, 'btn primary')
      ], 'dlg-actions')
    ];
  });
}

// ---------- chart: edit ----------
export function openChartDialog(el: any, api: { update: (patch: any) => void }): void {
  openDialog('Edit chart', (close) => {
    const draft: any = {
      chartType: el.chartType || 'barChart',
      title: el.title || '',
      legend: el.legend !== false,
      dataLabels: !!el.dataLabels,
      smooth: !!el.smooth,
      marker: !!el.marker,
      categories: Array.isArray(el.categories) ? el.categories.slice() : [],
      series: Array.isArray(el.series) ? JSON.parse(JSON.stringify(el.series)) : []
    };

    const typeSel = select(CHART_TYPES, draft.chartType);
    typeSel.addEventListener('change', () => {
      draft.chartType = typeSel.value;
    });
    const titleIn = input('text', String(draft.title ?? ''));
    titleIn.addEventListener('input', () => {
      draft.title = titleIn.value;
    });

    const catWrap = h('div', { class: 'dlg-col' });
    const serWrap = h('div', { class: 'dlg-col' });

    const renderSeries = () => {
      serWrap.innerHTML = '';
      draft.series.forEach((s: any, si: number) => {
        const nameIn = input('text', String(s.name ?? ''), 'dlg-cell');
        nameIn.addEventListener('input', () => {
          s.name = nameIn.value;
        });
        const vals = row([]);
        draft.categories.forEach((_: any, ci: number) => {
          const v = input('number', String((s.values || [])[ci] ?? 0), 'dlg-cell');
          v.addEventListener('input', () => {
            s.values = s.values || [];
            s.values[ci] = parseFloat(v.value) || 0;
          });
          vals.append(v);
        });
        serWrap.append(
          h('div', { class: 'dlg-block' }, [
            row([nameIn, btn('✕', () => { draft.series.splice(si, 1); renderSeries(); })]),
            vals
          ])
        );
      });
      serWrap.append(
        btn('+ Series', () => {
          draft.series.push({ name: 'Series ' + (draft.series.length + 1), values: draft.categories.map(() => 0) });
          renderSeries();
        })
      );
    };

    const renderCats = () => {
      catWrap.innerHTML = '';
      draft.categories.forEach((c: string, i: number) => {
        const inp = input('text', String(c ?? ''), 'dlg-cell');
        inp.addEventListener('input', () => {
          draft.categories[i] = inp.value;
        });
        catWrap.append(
          row([
            inp,
            btn('✕', () => {
              draft.categories.splice(i, 1);
              draft.series.forEach((s: any) => (s.values || []).splice(i, 1));
              renderCats();
              renderSeries();
            })
          ])
        );
      });
      catWrap.append(
        btn('+ Category', () => {
          draft.categories.push('');
          draft.series.forEach((s: any) => (s.values = s.values || []).push(0));
          renderCats();
          renderSeries();
        })
      );
    };

    renderCats();
    renderSeries();

    return [
      labelled('Type', typeSel),
      labelled('Title', titleIn),
      row([
        checkBox('Legend', draft.legend, (v) => (draft.legend = v)),
        checkBox('Data labels', draft.dataLabels, (v) => (draft.dataLabels = v)),
        checkBox('Smooth', draft.smooth, (v) => (draft.smooth = v)),
        checkBox('Markers', draft.marker, (v) => (draft.marker = v))
      ]),
      h('div', { class: 'dlg-split' }, [
        h('div', {}, [h('b', {}, ['Categories']), catWrap]),
        h('div', {}, [h('b', {}, ['Series']), serWrap])
      ]),
      row([
        btn('Cancel', close),
        btn('Save', () => {
          api.update(draft);
          close();
        }, 'btn primary')
      ], 'dlg-actions')
    ];
  });
}

// ---------- shapes ----------
export function openShapePicker(api: { addShape: (type: string) => void }): void {
  openDialog('Insert shape', (close) => {
    const wrap = h('div', { class: 'dlg-shapes' });
    for (const s of SHAPES as any[]) {
      const item = document.createElement('button');
      item.className = 'shape-item';
      item.title = s.name;
      item.innerHTML = `<svg viewBox="0 0 24 24" width="26" height="26" aria-hidden="true"><path d="${s.d}" fill="currentColor"/></svg><span>${s.name}</span>`;
      item.addEventListener('click', () => {
        api.addShape(s.type);
        close();
      });
      wrap.append(item);
    }
    return [wrap];
  });
}

// ---------- media ----------
export function openMediaDialog(api: {
  addMedia: (type: 'video' | 'audio', data: string, ext: string) => void;
}): void {
  openDialog('Insert media', (close) => {
    const pick = (type: 'video' | 'audio', accept: string) => {
      const file = input('file', '', '');
      file.style.display = 'none';
      file.accept = accept;
      file.addEventListener('change', () => {
        const f = file.files?.[0];
        if (!f) return;
        const rd = new FileReader();
        rd.onload = () => {
          api.addMedia(type, String(rd.result), (f.name.split('.').pop() || 'mp4').toLowerCase());
          close();
        };
        rd.readAsDataURL(f);
      });
      document.body.append(file);
      file.click();
    };
    return [
      row([
        btn('Video', () => pick('video', 'video/*,.mp4,.m4v,.mov,.webm,.avi')),
        btn('Audio', () => pick('audio', 'audio/*,.mp3,.m4a,.wav,.aac,.ogg,.wma'))
      ]),
      row([btn('Cancel', close)], 'dlg-actions')
    ];
  });
}

// ---------- keyboard shortcuts help ----------
export function openShortcuts(): void {
  const groups: { title: string; rows: [string, string][] }[] = [
    {
      title: 'General',
      rows: [
        ['F5', 'Start presentation'],
        ['Ctrl/Cmd + Z', 'Undo'],
        ['Ctrl/Cmd + Y / Shift+Z', 'Redo'],
        ['Ctrl/Cmd + A', 'Select all'],
        ['Ctrl/Cmd + C / X / V', 'Copy / Cut / Paste'],
        ['Esc', 'Cancel / deselect / exit group edit'],
        ['?', 'This help']
      ]
    },
    {
      title: 'Selection & arrange',
      rows: [
        ['Click', 'Select element'],
        ['Shift + Click', 'Add to selection'],
        ['Drag', 'Move'],
        ['Corner handles', 'Resize'],
        ['Top handle', 'Rotate'],
        ['Arrow keys', 'Nudge (Shift = ×10)'],
        ['Double-click group', 'Enter group editing']
      ]
    },
    {
      title: 'Text',
      rows: [
        ['Double-click text', 'Edit inline'],
        ['Ctrl/Cmd + B / I / U', 'Bold / Italic / Underline']
      ]
    }
  ];
  openDialog('Keyboard shortcuts', (close) => {
    const wrap = h('div', { class: 'shortcuts' });
    for (const g of groups) {
      wrap.append(h('h4', {}, [g.title]));
      const table = h('table', {});
      for (const [k, d] of g.rows) {
        table.append(h('tr', {}, [h('td', { class: 'key' }, [k]), h('td', {}, [d])]));
      }
      wrap.append(table);
    }
    return [wrap, h('div', { class: 'dlg-actions' }, [btn('Close', close, 'btn primary')])];
  });
}
