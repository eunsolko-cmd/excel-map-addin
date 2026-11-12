// 채번 작업창: 열린 통합문서의 「채번대장」 시트를 읽고 쓴다. 시트가 없으면 만든다.
(function () {
  const F = window.Fileno;
  const SHEET = '채번대장';
  const LAST_COL = String.fromCharCode(64 + F.HEADERS.length); // M
  const $ = id => document.getElementById(id);
  const state = { kind: '', office: '', lob: '' };
  let rows = [];

  function today() { const t = new Date(); return t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0') + '-' + String(t.getDate()).padStart(2, '0'); }
  function say(text, cls) { $('msg').textContent = text || ''; $('msg').className = cls || ''; }

  // 엑셀
  async function ensureSheet(ctx) {
    let ws = ctx.workbook.worksheets.getItemOrNullObject(SHEET);
    await ctx.sync();
    if (ws.isNullObject) {
      ws = ctx.workbook.worksheets.add(SHEET);
      const h = ws.getRange('A1:' + LAST_COL + '1');
      h.values = [F.HEADERS];
      h.format.font.bold = true;
      h.format.fill.color = '#DDEBF7';
      ws.freezePanes.freezeRows(1);
      const widths = [110, 60, 60, 170, 62, 70, 62, 40, 75, 85, 115, 55, 150];
      widths.forEach((w, i) => { ws.getRange(String.fromCharCode(65 + i) + ':' + String.fromCharCode(65 + i)).format.columnWidth = w; });
      await ctx.sync();
    }
    return ws;
  }
  async function readRows() {
    return Excel.run(async ctx => {
      const ws = await ensureSheet(ctx);
      const used = ws.getUsedRange();
      used.load('values');
      await ctx.sync();
      return used.values.slice(1).filter(r => r.some(v => v !== '' && v !== null));
    });
  }
  async function appendRow(row) {
    return Excel.run(async ctx => {
      const ws = await ensureSheet(ctx);
      const used = ws.getUsedRange();
      used.load('rowCount,rowIndex');
      await ctx.sync();
      const n = used.rowIndex + used.rowCount + 1; // 1부터 세는 다음 줄
      const rng = ws.getRange('A' + n + ':' + LAST_COL + n);
      rng.numberFormat = [F.HEADERS.map(h => (h === 'BKG(원)' ? '#,##0' : '@'))];
      rng.values = [row];
      ws.activate();
      rng.select();
      await ctx.sync();
      return n;
    });
  }
  async function cancelSelected() {
    return Excel.run(async ctx => {
      const sel = ctx.workbook.getSelectedRange();
      sel.load('rowIndex,rowCount');
      sel.worksheet.load('name');
      await ctx.sync();
      if (sel.worksheet.name !== SHEET) return { ok: false, error: '「' + SHEET + '」 시트에서 취소할 줄을 먼저 골라 주세요' };
      if (sel.rowCount !== 1 || sel.rowIndex < 1) return { ok: false, error: '취소할 줄 하나를 골라 주세요(머리줄 말고)' };
      const n = sel.rowIndex + 1;
      const line = sel.worksheet.getRange('A' + n + ':' + LAST_COL + n);
      line.load('values');
      await ctx.sync();
      const v = line.values[0];
      const num = v[F.COL['발급번호']];
      if (!F.parseNumber(num)) return { ok: false, error: '그 줄에는 발급번호가 없습니다' };
      if (v[F.COL['상태']] === '취소') return { ok: false, error: num + ' 는 이미 취소입니다' };
      sel.worksheet.getRange(String.fromCharCode(65 + F.COL['상태']) + n).values = [['취소']];
      await ctx.sync();
      return { ok: true, number: num };
    });
  }

  // 화면
  function input() {
    return {
      requester: $('requester').value.trim(), contract: $('contract').value.trim(), title: $('title').value.trim(),
      kind: state.kind, office: state.office, lob: state.lob, issueDate: $('issueDate').value,
      bkg: $('bkg').value, preChange: $('preChange').checked, note: $('note').value.trim(),
    };
  }
  function refreshPreview() {
    $('endtBox').style.display = state.kind === '배서' ? '' : 'none';
    const r = F.decide(input(), rows);
    const box = $('preview');
    box.className = r.ok ? (r.noNumber ? 'skip' : '') : 'err';
    $('pnum').textContent = r.ok ? (r.noNumber ? r.reason : r.number) : r.error;
    $('warns').innerHTML = (r.warnings || []).map(w => '· ' + w).join('<br>');
    $('go').disabled = !r.ok;
    $('go').textContent = r.ok && r.noNumber ? '발번 안 함으로 대장에 적기' : '번호 발급하고 대장에 적기';
  }
  function renderRecent() {
    const last = rows.slice(-8).reverse();
    $('recent').innerHTML = last.length ? last.map(r => {
      const x = r[F.COL['상태']] === '취소' ? ' class="x"' : '';
      const esc = s => String(s == null ? '' : s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
      return '<tr' + x + '><td class="n">' + esc(r[F.COL['발급번호']]) + '</td><td>' + esc(r[F.COL['계약번호']]) + ' ' + esc(r[F.COL['건 이름']]) +
        '<div class="hint">' + esc(r[F.COL['요청자']]) + ' · ' + esc(r[F.COL['발급일시']]) + (x ? ' · 취소' : '') + '</div></td></tr>';
    }).join('') : '<tr><td class="hint">아직 기록이 없습니다</td></tr>';
    const names = [...new Set(rows.map(r => r[F.COL['요청자']]).filter(Boolean))];
    $('reqlist').innerHTML = names.map(n => '<option value="' + String(n).replace(/"/g, '') + '">').join('');
  }
  async function reload() {
    try { rows = await readRows(); renderRecent(); refreshPreview(); }
    catch (e) { say('대장을 못 읽었습니다: ' + (e.message || e), 'bad'); }
  }
  function copyText(t) {
    const done = () => say(t + ' 복사했습니다', 'ok');
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(done, fallback); else fallback();
    function fallback() { const a = document.createElement('textarea'); a.value = t; document.body.appendChild(a); a.select(); document.execCommand('copy'); a.remove(); done(); }
  }

  async function issue() {
    $('go').disabled = true; say('');
    try {
      rows = await readRows(); // 쓰기 직전에 다시 읽어 번호가 겹치지 않게
      const inp = input();
      const r = F.decide(inp, rows);
      if (!r.ok) { say(r.error, 'bad'); refreshPreview(); return; }
      const line = await appendRow(F.buildRow(inp, r));
      rows = await readRows();
      const shown = r.noNumber ? '발번 안 함' : r.number;
      $('dnum').textContent = shown;
      $('copy').style.display = r.noNumber ? 'none' : '';
      $('done').style.display = 'block';
      say(line + '번째 줄에 적었습니다', 'ok');
      if (!r.noNumber) localStorage.setItem('fileno.lastReq', inp.requester);
      ['contract', 'title', 'bkg', 'note'].forEach(id => { $(id).value = ''; });
      $('preChange').checked = false;
      renderRecent(); refreshPreview();
    } catch (e) { say('적지 못했습니다: ' + (e.message || e), 'bad'); refreshPreview(); }
  }

  async function addExisting() {
    say('');
    try {
      rows = await readRows();
      const c = F.checkExisting($('exNum').value, rows);
      if (!c.ok) { say(c.error, 'bad'); return; }
      const m = c.prefix.replace(/^EN-/, '').match(/^(H[KSU])([PCMA])(\d{2})(\d{2})(\d{2})$/);
      const inp = {
        requester: $('exReq').value.trim(), contract: $('exContract').value.trim(), title: $('exTitle').value.trim(),
        kind: c.prefix.startsWith('EN-') ? '배서' : '', preChange: c.prefix.startsWith('EN-'),
        office: m[1], lob: m[2], issueDate: '20' + m[3] + '-' + m[4] + '-' + m[5], note: '옮겨 적음',
      };
      const row = F.buildRow(inp, { number: c.number });
      row[F.COL['상태']] = '발급';
      const line = await appendRow(row);
      ['exNum', 'exContract', 'exTitle'].forEach(id => { $(id).value = ''; });
      await reload();
      say(c.number + ' 를 ' + line + '번째 줄에 옮겨 적었습니다', 'ok');
    } catch (e) { say('적지 못했습니다: ' + (e.message || e), 'bad'); }
  }

  function wire() {
    document.querySelectorAll('.seg').forEach(seg => {
      seg.addEventListener('click', e => {
        const b = e.target.closest('button'); if (!b) return;
        state[seg.dataset.k] = b.dataset.v;
        seg.querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b));
        refreshPreview();
      });
    });
    ['requester', 'contract', 'title', 'bkg', 'issueDate', 'note'].forEach(id => $(id).addEventListener('input', () => { $('done').style.display = 'none'; refreshPreview(); }));
    $('preChange').addEventListener('change', refreshPreview);
    $('bkg').addEventListener('blur', () => { const n = Number($('bkg').value.replace(/[^\d]/g, '')); if (n) $('bkg').value = n.toLocaleString('en-US'); });
    $('go').addEventListener('click', issue);
    $('copy').addEventListener('click', () => copyText($('dnum').textContent));
    $('reload').addEventListener('click', reload);
    $('cancelSel').addEventListener('click', async () => {
      try { const r = await cancelSelected(); if (r.ok) { await reload(); say(r.number + ' 를 취소로 표시했습니다', 'ok'); } else say(r.error, 'bad'); }
      catch (e) { say('취소하지 못했습니다: ' + (e.message || e), 'bad'); }
    });
    $('exGo').addEventListener('click', addExisting);
    $('issueDate').value = today();
    try { $('requester').value = localStorage.getItem('fileno.lastReq') || ''; } catch (e) { /* 저장소 막힘 */ }
  }

  Office.onReady(() => { wire(); reload(); });
})();
