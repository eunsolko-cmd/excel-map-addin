/* 코리안리 송금신청 취합: 판정 규칙(브라우저 추가 기능과 node 시험이 같은 파일을 쓴다). */
(function (root) {
  'use strict';
  var SEP = /[\s,;\/()\[\]|\-:# ]+/;
  var CURS = ['KRW', 'USD', 'EUR', 'GBP', 'JPY', 'CNY'];
  var TOL = 10;

  function decodeEntities(s) {
    return s.replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"').replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(+d); })
      .replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCharCode(parseInt(h, 16)); });
  }
  function cellText(html) {
    return decodeEntities(html.replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
  }
  /* 메일 HTML 본문의 표를 줄(칸 글자 배열)로. 표가 없으면 글자 줄을 탭·두 칸 공백으로 나눈다. */
  function rowsFromMailHtml(html) {
    var out = [];
    var body = String(html || '').replace(/<!--[\s\S]*?-->/g, '').replace(/<(style|script)[\s\S]*?<\/\1>/gi, '');
    var trRe = /<tr[\s\S]*?<\/tr>/gi, m;
    while ((m = trRe.exec(body))) {
      var cells = [], tdRe = /<t[dh][\s\S]*?<\/t[dh]>/gi, c;
      while ((c = tdRe.exec(m[0]))) cells.push(cellText(c[0]));
      if (cells.some(function (x) { return x; })) out.push(cells);
    }
    if (!out.length) {
      cellText(body.replace(/<\/(p|div|tr)>/gi, '\n').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ''))
        .split('\n').forEach(function (l) { if (l.trim()) out.push(l.split(/\t| {2,}/)); });
    }
    return out;
  }
  function tokensOf(t) {
    if (!t) return [];
    return String(t).split(SEP).filter(function (x) {
      return x.length === 14 && (x.slice(0, 2) === 'U2' || x.slice(0, 2) === 'C2') && /^\d+$/.test(x.slice(1));
    });
  }
  function numOf(v) {
    var n;
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') n = v;
    else {
      var s = String(v).toUpperCase().replace(/\b(KRW|USD|EUR|GBP|JPY|CNY|DB|CR)\b/g, '').replace(/[,\s ]/g, '');
      if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
      n = parseFloat(s);
    }
    if (Math.round(n) === n && n >= 1900 && n <= 2100) return null;
    if (Math.abs(n) >= 1e11) return null;
    return Math.abs(n);
  }
  function curOf(t) {
    var u = String(t || '').toUpperCase();
    for (var i = 0; i < CURS.length; i++) if (u.indexOf(CURS[i]) >= 0) return CURS[i];
    return null;
  }
  /* 붙여넣은(또는 메일에서 읽은) 줄들 -> 신청 목록 [{no, amount, cur, line, src}] */
  function requestsFromRows(rows, src) {
    var reqs = [];
    rows.forEach(function (cells, i) {
      var toks = [], nums = [], cur = null;
      cells.forEach(function (c) {
        var tk = tokensOf(c);
        if (tk.length) toks = toks.concat(tk);
        else { var n = numOf(c); if (n !== null) nums.push(n); }
        if (!cur) cur = curOf(c);
      });
      toks = toks.filter(function (x, j) { return toks.indexOf(x) === j; });
      if (!toks.length) return;
      var amount = nums.length ? Math.max.apply(null, nums) : null;
      toks.forEach(function (no) { reqs.push({ no: no, amount: amount, cur: cur, line: cells.filter(Boolean).join(' | '), src: src || '', row: i + 1 }); });
    });
    return reqs;
  }
  /* 장부 시트(헤더 포함 2차원 배열) 여러 개 -> {header, rows: [obj]} */
  function ledgerFromSheets(arrays) {
    var header = null, rows = [];
    arrays.forEach(function (a) {
      if (!a || !a.length) return;
      var h = a[0].map(function (x) { return String(x === null || x === undefined ? '' : x).trim(); });
      if (!header) header = h;
      for (var i = 1; i < a.length; i++) {
        var o = {}, any = false;
        for (var j = 0; j < h.length; j++) { o[h[j]] = a[i][j] === undefined ? null : a[i][j]; if (o[h[j]] !== null && o[h[j]] !== '') any = true; }
        if (!any || (!o['계약번호'] && !o['U/W번호'])) continue;
        var d = o['Detail No.'];
        o._detailBase = d ? String(d).split('-')[0] : null;
        rows.push(o);
      }
    });
    return { header: header || [], rows: rows };
  }
  function absSum(rs) { var s = 0; rs.forEach(function (r) { s += +(r['미청산금액'] || 0); }); return Math.abs(s); }
  function groupBy(rs, keys) {
    var map = {}, order = [];
    rs.forEach(function (r) {
      var k = keys.map(function (x) { return String(r[x]); }).join('\u0001');
      if (!map[k]) { map[k] = []; order.push(k); }
      map[k].push(r);
    });
    return order.map(function (k) { return map[k]; });
  }
  function indexLedger(ledger) {
    var idx = {};
    ledger.rows.forEach(function (r) {
      ['계약번호', 'U/W번호', 'OUR_REFERENCE', '_detailBase'].forEach(function (c) {
        var v = r[c]; if (!v) return; v = String(v);
        (idx[v] = idx[v] || []).indexOf(r) < 0 && idx[v].push(r);
      });
    });
    return idx;
  }
  function judge(req, ledger, idx, groupsCache) {
    var rows = (idx[req.no] || []).slice();
    if (req.cur) { var r1 = rows.filter(function (r) { return r['화폐코드'] === req.cur; }); if (r1.length) rows = r1; }
    var want = req.amount;
    var close = function (t) { return want !== null && Math.abs(absSum(t) - want) <= TOL; };
    if (!rows.length) {
      if (want !== null) {
        if (!groupsCache.g) groupsCache.g = groupBy(ledger.rows, ['계약번호', '계수구분명', '적용년월']).map(function (t) { return { t: t, s: absSum(t) }; });
        var cand = groupsCache.g.filter(function (x) { return Math.abs(x.s - want) <= TOL && (!req.cur || x.t[0]['화폐코드'] === req.cur); });
        if (cand.length === 1) return { rows: cand[0].t, how: '번호가 장부에 없어 금액으로 찾음, 확인 필요' };
      }
      return { rows: [], how: '장부에 없음' };
    }
    if (want === null) return { rows: rows, how: '금액 못 읽음, 전부 가져옴' };
    if (close(rows)) return { rows: rows, how: '전부 맞음' };
    var kinds = groupBy(rows, ['계수구분명']);
    var byKind = kinds.filter(close);
    if (byKind.length === 1) return { rows: byKind[0], how: '계수구분 하나로 맞음' };
    var byYm = groupBy(rows, ['계수구분명', '적용년월']).filter(close);
    if (byYm.length === 1) return { rows: byYm[0], how: '계수구분+적용년월로 맞음' };
    var prefix = [];
    kinds.forEach(function (gt) {
      var months = gt.map(function (r) { return String(r['적용년월']); }).filter(function (x, i, a) { return a.indexOf(x) === i; }).sort();
      months.forEach(function (m) { var t = gt.filter(function (r) { return String(r['적용년월']) <= m; }); if (close(t)) prefix.push(t); });
    });
    if (prefix.length === 1) return { rows: prefix[0], how: '적용년월 어느 달까지로 맞음' };
    var singles = rows.filter(function (r) { return Math.abs(Math.abs(+(r['미청산금액'] || 0)) - want) <= TOL; });
    if (singles.length === 1) return { rows: singles, how: '한 줄로 맞음' };
    return { rows: rows, how: '금액 안 맞음, 확인 필요' };
  }
  /* 전체: 신청 목록 + 장부 -> {picked: [행 + 신청번호·판정], summary: [...]} */
  function consolidate(reqs, ledger) {
    var idx = indexLedger(ledger), cache = {}, picked = [], summary = [], seen = {};
    reqs.forEach(function (q) {
      var key = q.no + '|' + (q.cur || '') + '|' + q.amount;
      if (seen[key]) return; seen[key] = true;
      var j = judge(q, ledger, idx, cache);
      j.rows.forEach(function (r) { var o = {}; ledger.header.forEach(function (h) { o[h] = r[h]; }); o['신청번호'] = q.no; o['판정'] = j.how; o['담당 메일'] = q.src; picked.push(o); });
      var f = j.rows[0] || {};
      var sum = j.rows.length ? absSum(j.rows) : null;
      summary.push({ '담당 메일': q.src, '신청번호': q.no, '청산사코드': f['청산사코드'] || '', '계약자명': f['계약자명'] || '', '화폐': f['화폐코드'] || q.cur || '',
        '신청금액': q.amount, '장부합계': sum, '차이(장부-신청)': (sum === null || q.amount === null) ? null : Math.round((sum - q.amount) * 100) / 100,
        '줄수': j.rows.length, '판정': j.how, '메일 줄': q.line });
    });
    return { picked: picked, summary: summary, header: ledger.header.concat(['신청번호', '판정', '담당 메일']) };
  }
  /* ---------- Total List: 담당자 제출 파일에서 이번 주 변경 줄 모으기 ----------
     담당자 파일 = 회사 Total List 양식(머리줄 A칸 「파일번호」). 이번 주 줄은 필터로 남긴 «보이는 줄»,
     필터가 없으면 노란 칠(FFFF00)이 있는 줄로 본다. SheetJS 는 cellStyles:true 로 읽어야 숨김·칠이 보인다. */
  function totalListFromWorkbook(XLSX, wb, label, email) {
    var out = [];
    wb.SheetNames.forEach(function (sn) {
      var ws = wb.Sheets[sn]; if (!ws || !ws['!ref']) return;
      var a = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      var hi = -1;
      for (var i = 0; i < Math.min(a.length, 30); i++) { if (a[i] && String(a[i][0] || '').trim() === '파일번호') { hi = i; break; } }
      if (hi < 0) return;
      var rowsMeta = ws['!rows'] || [], anyHidden = false, range = XLSX.utils.decode_range(ws['!ref']);
      for (var r = hi + 1; r < a.length; r++) if (rowsMeta[r] && rowsMeta[r].hidden) { anyHidden = true; break; }
      var yellowRow = function (r) {
        for (var c = range.s.c; c <= range.e.c; c++) {
          var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
          var fg = cell && cell.s && cell.s.fgColor && cell.s.fgColor.rgb;
          if (fg && /FFFF00$/i.test(fg)) return true;
        }
        return false;
      };
      var header = a[hi].map(function (x, j) { var h = String(x === null || x === undefined ? '' : x).replace(/\s+/g, ' ').trim(); return h || ('칸' + (j + 1)); });
      var picked = [], dataRows = 0;
      for (var r2 = hi + 1; r2 < a.length; r2++) {
        var row = a[r2]; if (!row || row[0] === null || row[0] === undefined || String(row[0]).trim() === '') continue;
        if (!/\d/.test(String(row[0]))) continue;
        dataRows++;
        var take = anyHidden ? !(rowsMeta[r2] && rowsMeta[r2].hidden) : yellowRow(r2);
        if (take) picked.push(row);
      }
      out.push({ label: label, email: email || '', sheet: sn, header: header, rows: picked, dataRows: dataRows,
        how: anyHidden ? '필터로 남긴 줄' : '노란 칠 줄', asOf: (a[0] && typeof a[0][0] === 'number') ? a[0][0] : null });
    });
    return out;
  }
  /* 여러 담당자 결과 + 요청 메일 받는 사람 목록 -> 한 표 + 담당자별 건수 + 안 낸 사람 */
  function consolidateTotalList(parts, expected) {
    var header = null, rows = [], byWho = {}, keys = {};
    parts.forEach(function (p) {
      if (!header) header = p.header;
      var who = p.label;
      byWho[who] = byWho[who] || { 담당: who, 메일: p.email, 시트: p.sheet, 건수: 0, 전체행: p.dataRows, 방법: p.how };
      p.rows.forEach(function (r) {
        var o = { '담당(보낸 사람)': who };
        p.header.forEach(function (h, j) { o[h] = r[j] === undefined ? null : r[j]; });
        var k = String(r[0]) + '|' + String(r[1] === null || r[1] === undefined ? '' : r[1]);
        o['겹침'] = (keys[k] && keys[k] !== who) ? ('같은 파일번호·일련번호를 ' + keys[k] + '도 냄') : '';
        if (!keys[k]) keys[k] = who;
        rows.push(o);
        byWho[who].건수++;
      });
    });
    var sentEmails = parts.map(function (p) { return (p.email || '').toLowerCase(); });
    var missing = (expected || []).filter(function (e) { return e.email && sentEmails.indexOf(e.email.toLowerCase()) < 0; });
    return { header: ['담당(보낸 사람)'].concat(header || []).concat(['겹침']), rows: rows,
      who: Object.keys(byWho).map(function (k) { return byWho[k]; }), missing: missing };
  }
  var api = { rowsFromMailHtml: rowsFromMailHtml, requestsFromRows: requestsFromRows, ledgerFromSheets: ledgerFromSheets, consolidate: consolidate, tokensOf: tokensOf, numOf: numOf,
    totalListFromWorkbook: totalListFromWorkbook, consolidateTotalList: consolidateTotalList };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.RemitLogic = api;
})(this);
