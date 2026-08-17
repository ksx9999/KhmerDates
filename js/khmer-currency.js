/* ===== Khmer currency in words =====
 *
 * Ported from ReadKHR&USD.bas (KhmerCurrency / DollarsCurrency). The VBA is the
 * specification: where it does something surprising, this follows it rather
 * than "correcting" it, so numbers read the same as in the spreadsheets that
 * already use the macro.
 *
 * Two of those quirks are worth naming:
 *   - The `Case "One"` branches in the VBA are unreachable. GetHundreds returns
 *     Khmer text, never the literal string "One", so one riel reads through the
 *     general branch. Reproduced by simply not having those branches.
 *   - Riel centimes are suffixed រៀលគត់, not a centime word. That is what the
 *     macro does, so it is what this does.
 */
const KhmerCurrency = (() => {
  'use strict';

  // Place(1) is empty in the VBA (ReDim leaves it ""), and 6+ are never set,
  // so anything past a trillion loses its place word exactly as it does there.
  const PLACE = ['', '', 'ពាន់ ', 'លាន ', 'ពាន់លាន ', 'ពាន់ពាន់លាន ', '', '', '', ''];

  const DIGIT = ['', 'មួយ', 'ពីរ', 'បី', 'បួន', 'ប្រាំ',
                 'ប្រាំមួយ', 'ប្រាំពីរ', 'ប្រាំបី', 'ប្រាំបួន'];

  const TEEN = ['ដប់', 'ដប់មួយ', 'ដប់ពីរ', 'ដប់បី', 'ដប់បួន', 'ដប់ប្រាំ',
                'ដប់ប្រាំមួយ', 'ដប់ប្រាំពីរ', 'ដប់ប្រាំបី', 'ដប់ប្រាំបួន'];

  const TENS = ['', '', 'ម្ភៃ', 'សាមសិប', 'សែសិប', 'ហាសិប',
                'ហុកសិប', 'ចិតសិប', 'ប៉ែតសិប', 'កៅសិប'];

  const ZERO   = 'គ្មានលេខ';
  const RIEL   = 'រៀលគត់';
  const DOLLAR = 'ដុល្លារ';
  const EXACT  = 'គត់';
  const AND    = 'និង';
  const CENT   = 'សេន';

  function digit(ch) {
    const n = parseInt(ch, 10);
    return isNaN(n) ? '' : DIGIT[n];
  }

  /** Always receives two characters, as it does in the VBA. */
  function tens(text) {
    const t = parseInt(text[0], 10) || 0;
    if (t === 1) return TEEN[parseInt(text, 10) - 10] || '';
    return TENS[t] + digit(text[1]);
  }

  function hundreds(numStr) {
    if (parseInt(numStr, 10) === 0 || isNaN(parseInt(numStr, 10))) return '';
    const s = ('000' + numStr).slice(-3);
    let out = '';
    if (s[0] !== '0') out += digit(s[0]) + 'រយ';
    out += (s[1] !== '0') ? tens(s.slice(1)) : digit(s[2]);
    return out;
  }

  /** Splits the whole part into groups of three, right to left. */
  function groups(whole) {
    let out = '', count = 1, rest = whole;
    while (rest !== '') {
      const chunk = hundreds(rest.slice(-3));
      if (chunk !== '') out = chunk + PLACE[count] + out;
      rest = rest.length > 3 ? rest.slice(0, -3) : '';
      count++;
    }
    return out;
  }

  /** Normalises to the "1234.56" shape VBA's Str + Trim produces. */
  function split(value) {
    let s = String(value == null ? '' : value).trim().replace(/,/g, '');
    if (s === '' || isNaN(Number(s))) return null;
    s = String(Number(s));
    const dot = s.indexOf('.');
    if (dot < 0) return { whole: s, frac: null };
    return { whole: s.slice(0, dot), frac: (s.slice(dot + 1) + '00').slice(0, 2) };
  }

  function riel(value) {
    const p = split(value);
    if (!p) return '';
    let words = groups(p.whole);
    words = (words === '') ? ZERO : words + RIEL;
    let cents = p.frac ? tens(p.frac) : '';
    cents = cents ? AND + cents + RIEL : '';
    return (words + ' ' + cents).trim();
  }

  function dollar(value) {
    const p = split(value);
    if (!p) return '';
    let words = groups(p.whole);
    words = (words === '') ? ZERO : words + DOLLAR;
    // The VBA appends គត់ only when the source had no decimal point at all.
    if (p.frac === null) words += EXACT;
    let cents = p.frac ? tens(p.frac) : '';
    cents = cents ? AND + cents + CENT : '';
    return (words + ' ' + cents).trim();
  }

  return { riel, dollar };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = KhmerCurrency;
