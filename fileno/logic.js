// 채번 규칙. 브라우저(작업창)와 node(시험) 둘 다에서 쓴다.
// 근거: 10-29 안기혜 이사님 공지 파일번호표, 11-02 회의(은솔 전언), 2026-11-12 김가영 과장님 확인(앞자리+날짜마다 01부터), 은솔 「EN- 는 순번 따로」.
(function (root) {
  const OFFICES = { HK: 'HWS Korea 직접', HS: 'Howden Specialty Asia Pacific', HU: 'Howden Insurance Brokers Limited(런던)' };
  const LOBS = { P: '재물', C: '특종', M: '해상', A: '항공' };
  const SMALL_ENDT_LIMIT = 5000000;
  const HEADERS = ['발급일시', '요청자', '계약번호', '건 이름', '원계약/배서', '사명변경 전 계약', '출재 오피스', '종목', '발행일자', 'BKG(원)', '발급번호', '상태', '비고'];
  const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i]));
  const NUM_RE = /^((?:EN-)?H[KSU][PCMA]\d{6})(\d{2})$/;

  function pad(n) { return String(n).padStart(2, '0'); }

  // 엑셀 날짜 일련번호, Date, 'YYYY-MM-DD' 문자열 모두 받는다
  function toDate(v) {
    if (v instanceof Date) return v;
    if (typeof v === 'number' && v > 20000) return new Date(Math.round((v - 25569) * 86400000));
    const m = String(v || '').trim().match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return null;
  }
  function ymd(d) { return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function yymmdd(d) { return String(d.getUTCFullYear()).slice(2) + pad(d.getUTCMonth() + 1) + pad(d.getUTCDate()); }
  function toAmount(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(String(v).replace(/[,\s원]/g, ''));
    return isFinite(n) ? n : null;
  }

  function parseNumber(s) {
    const m = String(s || '').trim().toUpperCase().match(NUM_RE);
    return m ? { prefix: m[1], serial: +m[2] } : null;
  }

  // input: { kind:'원계약'|'배서', preChange:bool, office, lob, issueDate, bkg }
  // rows: 대장 값(머리줄 뺀 2차원 배열)
  function decide(input, rows) {
    const missing = [];
    if (!input.kind) missing.push('원계약/배서');
    if (!input.office) missing.push('출재 오피스');
    if (!input.lob) missing.push('종목');
    const d = toDate(input.issueDate);
    if (!d) missing.push('발행일자');
    const bkg = toAmount(input.bkg);
    if (input.kind === '배서' && bkg === null) missing.push('BKG 금액');
    if (missing.length) return { ok: false, error: missing.join(', ') + ' 을(를) 넣어 주세요' };
    if (input.office === 'HK' && input.lob === 'A') return { ok: false, error: 'HK 항공 번호는 번호표에 없습니다' };
    if (input.kind === '배서' && bkg <= SMALL_ENDT_LIMIT) {
      return { ok: true, noNumber: true, reason: 'BKG 500만원 이하 배서라 발번하지 않습니다' };
    }
    const prefix = (input.kind === '배서' && input.preChange ? 'EN-' : '') + input.office + input.lob + yymmdd(d);
    let max = 0;
    for (const r of rows || []) {
      const p = parseNumber(r[COL['발급번호']]);
      if (p && p.prefix === prefix && p.serial > max) max = p.serial; // 취소 번호도 센다(다시 안 씀)
    }
    if (max >= 99) return { ok: false, error: prefix + ' 는 오늘 99번까지 다 썼습니다' };
    const number = prefix + pad(max + 1);
    return { ok: true, number, prefix, serial: max + 1, issueDate: ymd(d), warnings: sameContractWarnings(input, rows) };
  }

  function sameContractWarnings(input, rows) {
    const c = String(input.contract || '').trim();
    if (!c) return [];
    const hits = (rows || []).filter(r => String(r[COL['계약번호']]).trim() === c && parseNumber(r[COL['발급번호']]) && r[COL['상태']] !== '취소');
    return hits.map(r => {
      const extra = [r[COL['원계약/배서']], r[COL['건 이름']]].filter(Boolean).join(', ');
      return '같은 계약번호로 이미 ' + r[COL['발급번호']] + (extra ? ' (' + extra + ')' : '') + ' 가 나갔습니다';
    });
  }

  // 대장에 넣을 한 줄
  function buildRow(input, result, now) {
    const t = now || new Date();
    const stamp = t.getFullYear() + '-' + pad(t.getMonth() + 1) + '-' + pad(t.getDate()) + ' ' + pad(t.getHours()) + ':' + pad(t.getMinutes());
    const d = toDate(input.issueDate);
    return [
      stamp,
      input.requester || '',
      input.contract || '',
      input.title || '',
      input.kind || '',
      input.kind === '배서' ? (input.preChange ? '예' : '아니오') : '',
      input.office || '',
      input.lob || '',
      d ? ymd(d) : '',
      input.kind === '배서' ? toAmount(input.bkg) : '',
      result.noNumber ? '발번 안 함' : result.number,
      result.noNumber ? '발번 안 함' : (input.status || '발급'),
      input.note || '',
    ];
  }

  // 이미 나간 번호 옮겨 적기
  function checkExisting(number, rows) {
    const p = parseNumber(number);
    if (!p) return { ok: false, error: '번호 모양이 다릅니다 (예: HUP26110301, EN-HUP26111201)' };
    const n = String(number).trim().toUpperCase();
    if ((rows || []).some(r => String(r[COL['발급번호']]).trim().toUpperCase() === n)) return { ok: false, error: n + ' 는 이미 대장에 있습니다' };
    return { ok: true, number: n, prefix: p.prefix, serial: p.serial };
  }

  const api = { OFFICES, LOBS, HEADERS, COL, SMALL_ENDT_LIMIT, decide, buildRow, parseNumber, checkExisting, toDate, ymd };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Fileno = api;
})(typeof window !== 'undefined' ? window : globalThis);
