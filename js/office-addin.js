/* ===== Office Add-in mode =====
 *
 * Loaded only when the page is opened with ?office=1, which is the URL the
 * add-in manifest points at. A normal browser or the Capacitor apps never load
 * this file, so nothing here can affect them.
 *
 * The calendar itself is untouched. The day-detail sheet already renders each
 * reading in a .detail-full-row carrying the plain text in data-copy for its
 * copy button, so this adds a second button to those rows that writes the same
 * text into the document instead of the clipboard.
 */
(function () {
  'use strict';

  var LABEL = { km: 'បញ្ចូល', en: 'Insert', zh: '插入' };
  var OK    = { km: 'បានបញ្ចូល', en: 'Inserted', zh: '已插入' };
  /* PowerPoint only accepts text when a text box is actually in edit mode, and
     Excel refuses when the selection is not a single writable range. Office
     reports both the same way, so the message names the likely cause. */
  var FAIL  = {
    km: 'មិនអាចបញ្ចូលបានទេ — សូមចុចលើកន្លែងសរសេរជាមុនសិន',
    en: 'Could not insert — click where the text should go first',
    zh: '无法插入 — 请先点击要插入的位置'
  };

  function lang() {
    try { return (window.I18n && I18n.getLang()) || 'en'; } catch (e) { return 'en'; }
  }
  function say(map) { return map[lang()] || map.en; }

  /* The app's own toast is private to its module, so this uses its own. */
  function flash(text, bad) {
    var el = document.getElementById('office-flash');
    if (!el) {
      el = document.createElement('div');
      el.id = 'office-flash';
      document.body.appendChild(el);
    }
    el.textContent = text;
    el.className = bad ? 'bad show' : 'show';
    clearTimeout(flash._t);
    flash._t = setTimeout(function () { el.className = ''; }, 2600);
  }

  function insert(text) {
    if (!window.Office || !Office.context || !Office.context.document) return;
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      function (result) {
        var failed = result.status === Office.AsyncResultStatus.Failed;
        flash(failed ? say(FAIL) : say(OK), failed);
      }
    );
  }


  /* ===== Place prefix =====
   *
   * Khmer letters open with the place before the date:
   *   ខេត្តព្រះសីហនុ, ថ្ងៃទី១៧ ខែសីហា ឆ្នាំ២០២៦
   *
   * khmer-calendar.js already builds exactly this in gDatesPro(), ported from
   * the VBA, but with the province hard-coded in ADH[3] — no good for anyone
   * outside Preah Sihanouk. So the place is a field here instead, saved per
   * machine, and the row is assembled from the Gregorian row the detail sheet
   * has already rendered. That avoids re-deriving the date and cannot drift
   * from what the rest of the sheet shows.
   */
  var PLACE_KEY = 'kh-cal-place';
  var PLACE_DEFAULT = 'ខេត្តព្រះសីហនុ';
  var PLACE_LABEL = { km: 'ទីកន្លែង', en: 'Place', zh: '地点' };

  function getPlace() {
    try {
      var v = localStorage.getItem(PLACE_KEY);
      return v === null ? PLACE_DEFAULT : v;
    } catch (e) { return PLACE_DEFAULT; }
  }
  function setPlace(v) {
    try { localStorage.setItem(PLACE_KEY, v); } catch (e) {}
  }

  /* A single field at the top of the pane. Saved on every keystroke: there is
     no Save button to forget, and the value is one short string. */
  function mountPlaceBar() {
    if (document.getElementById('office-place-bar')) return;
    var app = document.querySelector('.cal-app');
    if (!app) return;

    var bar = document.createElement('div');
    bar.id = 'office-place-bar';

    var label = document.createElement('label');
    label.setAttribute('for', 'office-place-input');
    label.textContent = PLACE_LABEL[lang()] || PLACE_LABEL.en;

    var input = document.createElement('input');
    input.type = 'text';
    input.id = 'office-place-input';
    input.value = getPlace();
    input.placeholder = PLACE_DEFAULT;
    input.setAttribute('autocomplete', 'off');
    input.addEventListener('input', function () {
      setPlace(input.value);
      decorate();           // keep the composed row in step with the field
    });

    bar.appendChild(label);
    bar.appendChild(input);
    app.insertBefore(bar, app.firstChild);
  }

  /* Builds "<place>, <Gregorian>" from the row the sheet already rendered. */
  function placeRowText() {
    var greg = document.querySelector('#cal-detail-content .detail-greg');
    if (!greg) return null;
    var row = greg.closest('.detail-full-row');
    var btn = row && row.querySelector('[data-copy]');
    var date = btn ? btn.getAttribute('data-copy') : greg.textContent.trim();
    if (!date) return null;
    var place = getPlace().trim();
    return place ? place + ', ' + date : date;
  }

  /* Add an insert button to any detail row that does not have one yet. The
     sheet is re-rendered on every day tap, so this runs from an observer
     rather than once at startup. */
  function decorate() {
    var rows = document.querySelectorAll('#cal-detail-content .detail-full-row');
    Array.prototype.forEach.call(rows, function (row) {
      if (row.querySelector('.office-insert-btn')) return;
      var src = row.querySelector('[data-copy]');
      if (!src) return;
      var text = src.getAttribute('data-copy');

      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'office-insert-btn';
      btn.title = say(LABEL);
      btn.setAttribute('aria-label', say(LABEL));
      btn.innerHTML =
        '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor"' +
        ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
        '<path d="M12 5v14"/><path d="M19 12l-7 7-7-7"/></svg>';
      btn.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        insert(text);
      });
      row.appendChild(btn);
    });

    mountPlaceRow();
  }

  /* The composed place+date row is not part of the app's own sheet, so it is
     added here and kept up to date as the field is typed in. */
  function mountPlaceRow() {
    var content = document.getElementById('cal-detail-content');
    if (!content) return;
    var text = placeRowText();
    var existing = document.getElementById('office-place-row');

    if (!text) { if (existing) existing.remove(); return; }
    if (existing) {
      existing.querySelector('.detail-full').textContent = text;
      existing.querySelector('.office-insert-btn').dataset.text = text;
      return;
    }

    var row = document.createElement('div');
    row.className = 'detail-full-row';
    row.id = 'office-place-row';

    var cell = document.createElement('div');
    cell.className = 'detail-full detail-selectable';
    cell.textContent = text;

    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'office-insert-btn';
    btn.title = say(LABEL);
    btn.setAttribute('aria-label', say(LABEL));
    btn.dataset.text = text;
    btn.innerHTML =
      '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor"' +
      ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M12 5v14"/><path d="M19 12l-7 7-7-7"/></svg>';
    btn.addEventListener('click', function (e) {
      e.preventDefault();
      e.stopPropagation();
      insert(btn.dataset.text);
    });

    row.appendChild(cell);
    row.appendChild(btn);
    content.appendChild(row);
  }

  if (!window.Office || !Office.onReady) return;

  /*
   * Ask Office to reopen this pane by itself the next time the document is
   * opened, so the calendar is simply there rather than something to go and
   * find on the ribbon every time.
   *
   * This is per document, not per app: Office stores the flag inside the file,
   * and there is no equivalent that fires on application start for Word, Excel
   * or PowerPoint. So it takes effect from the next open of any document where
   * the calendar has been used and saved. It also only works because the
   * manifest's TaskpaneId is exactly Office.AutoShowTaskpaneWithDocument.
   */
  function autoOpenWithDocument() {
    try {
      var settings = Office.context && Office.context.document && Office.context.document.settings;
      if (!settings) return;
      if (settings.get('Office.AutoShowTaskpaneWithDocument') === true) return;
      settings.set('Office.AutoShowTaskpaneWithDocument', true);
      // Saving can fail harmlessly — a read-only or unsaved document, say — so
      // the callback deliberately does nothing rather than alarming the user.
      settings.saveAsync(function () {});
    } catch (e) { /* never let this stop the calendar rendering */ }
  }

  Office.onReady(function (info) {
    var host = (info && info.host ? String(info.host) : 'office').toLowerCase();
    document.body.classList.add('office-addin', 'office-host-' + host);
    autoOpenWithDocument();
    mountPlaceBar();

    var target = document.getElementById('cal-detail-content');
    if (target && window.MutationObserver) {
      new MutationObserver(decorate).observe(target, { childList: true, subtree: true });
    }
    decorate();
  });
})();
