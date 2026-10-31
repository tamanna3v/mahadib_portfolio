// ── VISITOR ANALYTICS ───────────────────────────────────
// Google Analytics 4 records who visits, for how long, from which source
// (Google, LinkedIn, Facebook, WhatsApp, direct…) and which pages they view.
// The reports live at analytics.google.com behind the owner's Gmail login.
// Page changes inside the single-page site are picked up from browser history
// by GA4's enhanced measurement, so every tab and blog post is logged.
// Paste the GA4 Measurement ID (G-XXXXXXXXXX) here to switch tracking on.
(function(){
  const GA_ID = 'G-11B0HNP710';
  window.trackEvent = function(name, params){
    if(window.gtag) window.gtag('event', name, params || {});
  };
  if(!GA_ID || location.hostname === 'localhost' || location.hostname === '127.0.0.1') return;

  window.dataLayer = window.dataLayer || [];
  window.gtag = function(){ window.dataLayer.push(arguments); };
  window.gtag('js', new Date());
  window.gtag('config', GA_ID);

  const s = document.createElement('script');
  s.async = true;
  s.src = 'https://www.googletagmanager.com/gtag/js?id=' + GA_ID;
  document.head.appendChild(s);
})();
