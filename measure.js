/* מדידה והסכמה — לעמודים המשניים של האתר.
   עמוד התפריט עצמו מחזיק את אותה לוגיקה בתוכו, ומשתמש באותו מפתח אחסון,
   כך שהבחירה של המבקר נשמרת בין כל עמודי האתר.

   Cloudflare נטען תמיד: הוא אינו יוצר עוגיות ואינו מזהה מבקרים.
   Google Analytics נטען רק אחרי אישור מפורש. */
(function () {
  'use strict';

  var MEASURE = {
    cloudflare: 'ffc9122d2c14416cb735e79f02d65d76',
    ga4: 'G-4K6WW34FKZ'
  };
  var KEY = 'migdalor_consent_v1';

  function state() {
    try { return localStorage.getItem(KEY) || ''; } catch (err) { return 'denied'; }
  }
  function save(v) {
    try { localStorage.setItem(KEY, v); } catch (err) {}
  }

  var gaLoaded = false;
  function loadGA() {
    if (gaLoaded || !MEASURE.ga4) return;
    gaLoaded = true;
    window.dataLayer = window.dataLayer || [];
    window.gtag = function () { window.dataLayer.push(arguments); };
    gtag('js', new Date());
    gtag('config', MEASURE.ga4, { send_page_view: false, anonymize_ip: true });
    var s = document.createElement('script');
    s.async = true;
    s.src = 'https://www.googletagmanager.com/gtag/js?id=' + encodeURIComponent(MEASURE.ga4);
    document.head.appendChild(s);
    gtag('event', 'page_view', { page_title: document.title, page_location: location.href });
  }

  // חזרה מאישור: מכבים את השליחה ומוחקים את העוגיות שגוגל כבר כתב
  function denyGA() {
    if (MEASURE.ga4) window['ga-disable-' + MEASURE.ga4] = true;
    var names = document.cookie.split(';')
      .map(function (c) { return c.split('=')[0].trim(); })
      .filter(function (n) { return /^_ga/.test(n) || /^_gid$/.test(n) || /^_gat/.test(n); });
    if (!names.length) return;
    var host = location.hostname, parts = host.split('.');
    var domains = ['', host, '.' + host];
    if (parts.length > 2) domains.push('.' + parts.slice(-2).join('.'));
    names.forEach(function (n) {
      domains.forEach(function (d) {
        document.cookie = n + '=; Max-Age=0; path=/' + (d ? '; domain=' + d : '');
      });
    });
  }

  function loadCloudflare() {
    if (!MEASURE.cloudflare) return;
    var b = document.createElement('script');
    b.defer = true;
    b.src = 'https://static.cloudflareinsights.com/beacon.min.js';
    b.setAttribute('data-cf-beacon', JSON.stringify({ token: MEASURE.cloudflare }));
    document.head.appendChild(b);
  }

  // הבאנר נבנה כאן ולא בכל עמוד בנפרד, כדי שהנוסח יישאר אחד
  function buildBar() {
    var bar = document.createElement('div');
    bar.className = 'consent-bar';
    bar.setAttribute('role', 'region');
    bar.setAttribute('aria-label', 'בחירת הסכמה למדידה');
    bar.innerHTML =
      '<span>אנחנו מודדים את השימוש באתר כדי לשפר אותו, בעזרת Google Analytics שמשתמש בעוגיות. ' +
      'אפשר לדחות — האתר יעבוד בדיוק אותו דבר. ' +
      '<a href="privacy.html">להצהרת הפרטיות המלאה</a></span>' +
      '<span class="consent-actions">' +
      '<button type="button" data-consent="denied" class="consent-ghost">דחייה</button>' +
      '<button type="button" data-consent="granted">אישור</button>' +
      '</span>';
    bar.addEventListener('click', function (e) {
      var btn = e.target.closest('[data-consent]');
      if (!btn) return;
      var v = btn.dataset.consent;
      save(v);
      bar.remove();
      if (v === 'granted') loadGA(); else denyGA();
    });
    document.body.appendChild(bar);
  }

  function start() {
    loadCloudflare();
    var v = state();
    if (v === 'granted') loadGA();
    else if (v === 'denied') denyGA();
    else buildBar();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
})();
