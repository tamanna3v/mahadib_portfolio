// ── LANGUAGE SWITCHER ───────────────────────────────────
// English, Bangla, French, Arabic, Portuguese.
// Navigation, section panel, footer and other site chrome use the hand-written
// strings below; the long-form page content (rendered by JS) is translated on
// the fly with Google Translate, which also follows SPA page changes.
(function(){
  const LANGS = [
    {code:'en', label:'English',   short:'EN'},
    {code:'bn', label:'বাংলা',      short:'বাং'},
    {code:'fr', label:'Français',  short:'FR'},
    {code:'ar', label:'العربية',    short:'ع', rtl:true},
    {code:'pt', label:'Português', short:'PT'}
  ];
  const CODES = LANGS.map(l=>l.code);
  const KEY = 'site-lang';

  // English source string → {bn, fr, ar, pt}
  const UI = {
    "Products":{bn:"প্রোডাক্টস",fr:"Produits",ar:"المنتجات",pt:"Produtos"},
    "AI Accelerator":{bn:"এআই অ্যাক্সিলারেটর",fr:"Accélérateur IA",ar:"مسرّع الذكاء الاصطناعي",pt:"Acelerador de IA"},
    "AI Briefing Session (AIBS)":{bn:"এআই ব্রিফিং সেশন (AIBS)",fr:"Session d'initiation à l'IA (AIBS)",ar:"جلسة إحاطة الذكاء الاصطناعي (AIBS)",pt:"Sessão de Briefing de IA (AIBS)"},
    "AIPA Session Plan":{bn:"AIPA সেশন পরিকল্পনা",fr:"Plan des sessions AIPA",ar:"خطة جلسات AIPA",pt:"Plano de Sessões AIPA"},
    "AI Strategy Accelerator (AISA)":{bn:"এআই স্ট্র্যাটেজি অ্যাক্সিলারেটর (AISA)",fr:"Accélérateur de stratégie IA (AISA)",ar:"مسرّع استراتيجية الذكاء الاصطناعي (AISA)",pt:"Acelerador de Estratégia de IA (AISA)"},
    "AI Inference":{bn:"এআই ইনফারেন্স",fr:"Inférence IA",ar:"استدلال الذكاء الاصطناعي",pt:"Inferência de IA"},
    "Consulting":{bn:"কনসাল্টিং",fr:"Conseil",ar:"الاستشارات",pt:"Consultoria"},
    "Research":{bn:"গবেষণা",fr:"Recherche",ar:"الأبحاث",pt:"Pesquisa"},
    "Events and News":{bn:"ইভেন্ট ও সংবাদ",fr:"Événements et actualités",ar:"الفعاليات والأخبار",pt:"Eventos e Notícias"},
    "Blog":{bn:"ব্লগ",fr:"Blog",ar:"المدونة",pt:"Blog"},
    "Contact":{bn:"যোগাযোগ",fr:"Contact",ar:"تواصل",pt:"Contato"},
    "Partner / Build Together":{bn:"পার্টনার / একসাথে গড়ি",fr:"Partenaire / Construire ensemble",ar:"شريك / لنبنِ معًا",pt:"Parceiro / Construir Juntos"},
    "Affiliation":{bn:"অ্যাফিলিয়েশন",fr:"Affiliation",ar:"برنامج الشراكة",pt:"Afiliação"},
    "Download Profile (PPT) ⬇":{bn:"প্রোফাইল ডাউনলোড (PPT) ⬇",fr:"Télécharger le profil (PPT) ⬇",ar:"تنزيل الملف التعريفي (PPT) ⬇",pt:"Baixar Perfil (PPT) ⬇"},
    "Sections":{bn:"অংশসমূহ",fr:"Sections",ar:"الأقسام",pt:"Seções"},
    "Intro":{bn:"পরিচিতি",fr:"Présentation",ar:"مقدمة",pt:"Introdução"},
    "Building":{bn:"যা গড়ছি",fr:"En construction",ar:"ما نبنيه",pt:"Construindo"},
    "Timeline":{bn:"টাইমলাইন",fr:"Parcours",ar:"المسيرة",pt:"Linha do Tempo"},
    "Awards":{bn:"পুরস্কার",fr:"Distinctions",ar:"الجوائز",pt:"Prêmios"},
    "Experience":{bn:"অভিজ্ঞতা",fr:"Expérience",ar:"الخبرة",pt:"Experiência"},
    "Building NOW & Companies":{bn:"এখন যা গড়ছি ও প্রতিষ্ঠানসমূহ",fr:"Projets actuels & entreprises",ar:"مشاريع حالية وشركات",pt:"Projetos atuais e empresas"},
    "Career · 20 yrs":{bn:"ক্যারিয়ার · ২০ বছর",fr:"Carrière · 20 ans",ar:"المسيرة المهنية · 20 عامًا",pt:"Carreira · 20 anos"},
    "Awards & Recognition":{bn:"পুরস্কার ও স্বীকৃতি",fr:"Prix & reconnaissance",ar:"الجوائز والتقدير",pt:"Prêmios e Reconhecimento"},
    "Work Experience & Skills":{bn:"কর্মঅভিজ্ঞতা ও দক্ষতা",fr:"Expérience & compétences",ar:"الخبرة العملية والمهارات",pt:"Experiência e Competências"},
    "© 2026 Hadi Bhuiyan · Dhaka, Bangladesh":{bn:"© ২০২৬ হাদি ভূঁইয়া · ঢাকা, বাংলাদেশ",fr:"© 2026 Hadi Bhuiyan · Dacca, Bangladesh",ar:"© 2026 هادي بويان · دكا، بنغلاديش",pt:"© 2026 Hadi Bhuiyan · Daca, Bangladesh"},
    "Chat on WhatsApp · EasyFlow":{bn:"হোয়াটসঅ্যাপে চ্যাট করুন · EasyFlow",fr:"Discuter sur WhatsApp · EasyFlow",ar:"تحدث عبر واتساب · EasyFlow",pt:"Conversar no WhatsApp · EasyFlow"}
  };
  const CHROME = 'nav, #sections-panel, footer, #wa-chat-bar';

  function readCookie(){
    const m = document.cookie.match(/(?:^|;\s*)googtrans=\/[^/]*\/([^;]+)/);
    return m ? decodeURIComponent(m[1]) : null;
  }
  function getLang(){
    const q = new URLSearchParams(location.search).get('lang');
    if(q && CODES.includes(q)) return q;
    let saved = null;
    try{ saved = localStorage.getItem(KEY); }catch(e){}
    if(saved && CODES.includes(saved)) return saved;
    const c = readCookie();
    return c && CODES.includes(c) ? c : 'en';
  }
  // Google Translate reads its target language from the googtrans cookie;
  // it has to be set (or cleared) on both the host and the parent domain.
  function setCookie(lang){
    const host = location.hostname;
    const parts = host.split('.');
    const domains = ['', host];
    if(parts.length > 2) domains.push('.' + parts.slice(-2).join('.'));
    domains.forEach(d=>{
      const dom = d ? ';domain=' + d : '';
      document.cookie = lang === 'en'
        ? 'googtrans=;path=/;expires=Thu, 01 Jan 1970 00:00:00 GMT' + dom
        : 'googtrans=/auto/' + lang + ';path=/' + dom;
    });
  }

  const lang = getLang();
  const meta = LANGS.find(l=>l.code===lang);
  setCookie(lang);
  try{ localStorage.setItem(KEY, lang); }catch(e){}
  document.documentElement.lang = lang;
  document.documentElement.dir = meta.rtl ? 'rtl' : 'ltr';
  document.documentElement.classList.add('lang-' + lang);

  function choose(code){
    if(code === lang) return;
    try{ localStorage.setItem(KEY, code); }catch(e){}
    setCookie(code);
    const u = new URL(location.href);
    u.searchParams.delete('lang');
    location.replace(u.toString());
  }

  // Swap exact-match chrome strings for the hand-written translation.
  function translateChrome(root){
    if(lang === 'en') return;
    (root || document).querySelectorAll(CHROME).forEach(scope=>{
      const walker = document.createTreeWalker(scope, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while(walker.nextNode()) nodes.push(walker.currentNode);
      nodes.forEach(n=>{
        const src = n.nodeValue.trim();
        const t = UI[src] && UI[src][lang];
        if(!t) return;
        n.nodeValue = n.nodeValue.replace(src, t);
        n.parentElement.classList.add('notranslate');
      });
    });
  }

  function buildSwitcher(){
    const wrap = document.createElement('div');
    wrap.className = 'lang-switch notranslate';
    wrap.setAttribute('translate', 'no');
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'lang-switch-btn';
    btn.setAttribute('aria-haspopup', 'listbox');
    btn.setAttribute('aria-label', 'Language');
    btn.innerHTML = '<span aria-hidden="true">🌐</span> ' + meta.short + ' <span class="lang-switch-arrow">▾</span>';
    const menu = document.createElement('div');
    menu.className = 'lang-switch-menu';
    menu.setAttribute('role', 'listbox');
    LANGS.forEach(l=>{
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'lang-switch-item' + (l.code === lang ? ' active' : '');
      item.setAttribute('role', 'option');
      item.setAttribute('lang', l.code);
      item.dir = l.rtl ? 'rtl' : 'ltr';
      item.textContent = l.label;
      item.onclick = ()=>choose(l.code);
      menu.appendChild(item);
    });
    btn.onclick = (e)=>{ e.stopPropagation(); wrap.classList.toggle('open'); };
    document.addEventListener('click', ()=>wrap.classList.remove('open'));
    wrap.appendChild(btn);
    wrap.appendChild(menu);
    const nav = document.querySelector('body > nav, nav.topnav');
    if(nav && nav.matches('body > nav')) nav.appendChild(wrap);
    else if(nav && nav.parentElement) nav.parentElement.insertBefore(wrap, nav.nextSibling);
    else { wrap.classList.add('lang-switch-float'); document.body.appendChild(wrap); }
  }

  function injectStyles(){
    const css = `
.lang-switch{position:relative;display:inline-block;flex:0 0 auto;font-family:inherit}
.lang-switch-btn{background:transparent;border:1px solid rgba(0,0,0,.18);color:inherit;padding:8px 12px;font-family:inherit;font-size:12px;font-weight:500;line-height:1;letter-spacing:.06em;cursor:pointer;white-space:nowrap}
.lang-switch-btn:hover{border-color:currentColor}
.lang-switch-arrow{font-size:9px;opacity:.7}
.lang-switch-menu{display:none;position:absolute;top:100%;right:0;margin-top:4px;background:#fff;border:1px solid rgba(0,0,0,.12);min-width:150px;padding:6px 0;z-index:10000;box-shadow:0 8px 24px rgba(0,0,0,.08)}
html[dir=rtl] .lang-switch-menu{right:auto;left:0}
.lang-switch.open .lang-switch-menu{display:block}
.lang-switch-item{display:block;width:100%;background:transparent;border:none;padding:9px 16px;font-size:14px;color:#1c2a21;cursor:pointer;text-align:start}
.lang-switch-item:hover{background:#f2f0ea}
.lang-switch-item.active{font-weight:600;color:#2e5c3f}
.lang-switch-float{position:fixed;bottom:24px;left:20px;z-index:9999;background:#fff}
html[dir=rtl] .lang-switch-float{left:auto;right:20px}
/* Hide Google Translate's own toolbar, tooltips and highlights */
body > .skiptranslate, iframe.skiptranslate, .goog-te-banner-frame, #goog-gt-tt, .goog-te-balloon-frame, .goog-tooltip, .VIpgJd-ZVi9od-ORHb-OEVmcd, .VIpgJd-ZVi9od-aZ2wEe-wOHMyf{display:none!important}
body{top:0!important}
font.goog-text-highlight, .VIpgJd-yAWNEb-VIpgJd-fmcmS-sn54Q{background:none!important;box-shadow:none!important}
#gt-hidden{display:none}
html.lang-bn body{font-family:'Hind Siliguri','Noto Sans Bengali',var(--sans, sans-serif)}
html.lang-ar body{font-family:'Noto Naskh Arabic','Segoe UI','Tahoma',var(--sans, sans-serif)}
/* Letter-spacing breaks Arabic letter joining; fixed side widgets mirror in RTL */
html.lang-ar body *{letter-spacing:normal!important}
html[dir=rtl] #sections-panel{right:auto;left:20px}
html[dir=rtl] #wa-chat-bar{right:auto;left:20px}
@media(max-width:768px){.lang-switch-btn{padding:7px 10px}}`;
    const s = document.createElement('style');
    s.textContent = css;
    document.head.appendChild(s);
    if(lang === 'ar' || lang === 'bn'){
      const f = document.createElement('link');
      f.rel = 'stylesheet';
      f.href = lang === 'ar'
        ? 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;600;700&display=swap'
        : 'https://fonts.googleapis.com/css2?family=Hind+Siliguri:wght@400;500;600&display=swap';
      document.head.appendChild(f);
    }
  }

  function loadGoogleTranslate(){
    if(lang === 'en') return;
    const holder = document.createElement('div');
    holder.id = 'gt-hidden';
    document.body.appendChild(holder);
    window.__gtInit = function(){
      new google.translate.TranslateElement({pageLanguage:'auto', autoDisplay:false}, 'gt-hidden');
    };
    const s = document.createElement('script');
    s.src = 'https://translate.google.com/translate_a/element.js?cb=__gtInit';
    s.async = true;
    document.body.appendChild(s);
  }

  function start(){
    injectStyles();
    buildSwitcher();
    translateChrome();
    loadGoogleTranslate();
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();

  window.SiteI18n = {lang, langs:LANGS, choose};
})();
