/* =====================================================================
   generate.js — يقرأ ملف الإكسل ويولّد صفحات HTML فيزيائية لكل منتج
   التشغيل:  node generate.js
   ===================================================================== */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

/* ============ الإعدادات — عدّلها من هنا ============ */
const CONFIG = {
  excelFile:    'amazon_details_v5_24.xlsx',  // اسم ملف الإكسل
  sheetName:    '',                            // اتركه فارغاً = أول ورقة
  outputDir:    'product',                     // مجلد صفحات المنتجات
  siteUrl:      'https://ishopping.fit',         // ⚠️ ضع رابط موقعك الحقيقي (مهم للـ SEO)
  siteName:     'سوقي',
  siteTagline:  'أفضل المنتجات بأفضل الأسعار',
  currency:     'EGP',
  currencyAr:   'جنيه',
  locale:       'ar_EG',
  // صفحة المتجر الديناميكية (تُربط من صفحات المنتجات)
  mainSiteUrl:  'https://ishopping.fit/index2.html',
  watchMode:    process.argv.includes('--watch'),
  watchIntervalMs: 30000                       // في وضع watch: تحقق كل 30 ثانية
};

/* ============ أدوات مساعدة ============ */
const IMG_RE = /https?:\/\/[^\s,|"'<>]+\.(?:jpg|jpeg|png|webp|gif)(?:\?[^\s,|"'<>]*)?/gi;

function safeStr(v){ return v === null || v === undefined ? '' : String(v).trim(); }
function escapeHTML(s){ return safeStr(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
function escapeAttr(s){ return safeStr(s).replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
function slugify(s){ return safeStr(s).toLowerCase().replace(/[^\w\u0600-\u06FF\-]+/g,'-').replace(/-+/g,'-').replace(/^-|-$/g,'').slice(0,80); }

function parseNum(v){
  if(v === null || v === undefined || v === '') return 0;
  let s = String(v).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٬،,\s]/g,'');
  const m = s.match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : 0;
}
function parseRating(v){
  if(!v) return 0;
  const m = String(v).match(/(\d+(?:[.,]\d+)?)/);
  if(!m) return 0;
  let r = parseFloat(m[1].replace(',','.'));
  if(r > 5) r = r/10;
  return Math.min(5, Math.max(0, r));
}
function parseReviews(v){
  if(!v) return 0;
  const m = String(v).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).match(/\d[\d,\.]*/);
  return m ? parseNum(m[0]) : 0;
}
function hiRes(url){
  if(!url) return '';
  return url.replace(/_AC_US40_/g,'_AC_SL800_').replace(/_AC_SR38,50_/g,'_AC_SL800_')
            .replace(/\.SS40_BG85,85,85_BR-120_PKdp-play-icon-overlay__\./g,'._AC_SL800_.');
}
function collectImages(row){
  const urls = [];
  for(const cell of row){
    if(cell === null || cell === undefined) continue;
    const matches = String(cell).match(IMG_RE);
    if(matches) for(const u of matches){
      IMG_RE.lastIndex = 0;
      if(!/link\.amazon|amzn\.to/i.test(u) && !urls.includes(u)) urls.push(u);
    }
  }
  return urls;
}
function findCol(headers, ...candidates){
  if(!headers || !headers.length) return -1;
  const clean = headers.map(h => safeStr(h).toLowerCase().trim());
  for(const c of candidates){
    const i = clean.indexOf(c.toLowerCase().trim());
    if(i !== -1) return i;
  }
  for(const c of candidates){
    const cc = c.toLowerCase().trim();
    const i = clean.findIndex(h => h && h.includes(cc));
    if(i !== -1) return i;
  }
  return -1;
}
function guessCategory(text){
  const t = safeStr(text).toLowerCase();
  if(/موبايل|هاتف|جوال|جالاكسي|ايفون|آيفون|اوبو|اونر|هونر|smartphone|phone|galaxy|iphone|honor|oppo/i.test(t)) return 'هواتف';
  if(/غلاية|مكواة|خلاط|مفرمة|محضر|لانش بوكس|كوب|زجاجة|كبة|kettle|iron|blender|mixer/i.test(t)) return 'أجهزة منزلية';
  if(/مجفف شعر|ماكينة حلاقة|شاحن|hair dryer|shaver|charger|power bank/i.test(t)) return 'إلكترونيات';
  if(/شامبو|كريم|عطر|عناية|beauty|skin|shampoo/i.test(t)) return 'العناية الشخصية';
  return 'عام';
}

/* ============ قراءة الإكسل ============ */
function readExcel(){
  const wb = XLSX.readFile(CONFIG.excelFile);
  const sheetName = CONFIG.sheetName && wb.SheetNames.includes(CONFIG.sheetName) ? CONFIG.sheetName : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', blankrows:false });
  if(!rows.length) throw new Error('الملف فارغ');

  const headers = rows[0].map(h => safeStr(h));
  const asinCol = findCol(headers, 'asin');
  const products = [];

  for(let i=1;i<rows.length;i++){
    const row = rows[i];
    if(!row || !row.length) continue;
    if(!row.some(c => safeStr(c) !== '')) continue;
    if(asinCol !== -1 && (!safeStr(row[asinCol]) || /^asin$/i.test(safeStr(row[asinCol])))) continue;

    const col = (...names) => {
      const idx = findCol(headers, ...names);
      return idx === -1 ? '' : safeStr(row[idx]);
    };

    const asin = col('asin') || ('ROW_' + i);
    const title = col('title_from_file','title','اسم المنتج','العنوان') || 'منتج بدون عنوان';
    const brand = col('brand','manufacturer','العلامة التجارية') || '';
    const price = parseNum(col('price','السعر'));
    const oldPrice = parseNum(col('price_original','السعر الأصلي'));
    let discount = parseNum(col('discount_percent','نسبة الخصم'));
    if(!discount && oldPrice > price && oldPrice > 0) discount = Math.round((1 - price/oldPrice) * 100);

    const rating = parseRating(col('rating','التقييم'));
    const reviews = parseReviews(col('reviews_count','عدد التقييمات'));
    const images = collectImages(row);
    const link = col('link','الرابط');
    const description = col('description','الوصف');
    const highlightsRaw = col('highlights','المميزات');
    const availability = col('availability','التوفر') || 'متوفر';
    const returns = col('returns','الإرجاع') || '';
    const delivery = col('delivery','التوصيل') || '';
    const seller = col('seller','البائع') || '';
    const highlights = highlightsRaw ? highlightsRaw.split(/\s*[|\n•·]\s*/).map(s=>s.trim()).filter(Boolean) : [];
    const category = guessCategory(title + ' ' + description + ' ' + brand);

    // كل الأعمدة كمواصفات
    const specs = [];
    for(let j=0;j<headers.length;j++){
      const h = safeStr(headers[j]); const v = safeStr(row[j]);
      if(!h || !v) continue;
      if(h.toLowerCase() === 'images') continue;
      IMG_RE.lastIndex = 0;
      if(IMG_RE.test(v)) continue;
      specs.push([h, v]);
    }

    products.push({
      asin, slug: slugify(asin), title, brand, price, oldPrice, discount,
      rating, reviews, images, image: images[0] || '', link, description,
      highlights, availability, returns, delivery, seller, category, specs
    });
  }
  return products;
}

/* ============ توليد صفحة منتج ============ */
function generateProductPage(p){
  const title = p.title;
  const brand = p.brand || CONFIG.siteName;
  const img = hiRes(p.image) || '';
  const descShort = (p.description || '').replace(/\s+/g,' ').slice(0, 158);
  const pageUrl = `${CONFIG.siteUrl.replace(/\/$/,'')}/${CONFIG.outputDir}/${p.asin}.html`;
  const keywords = [title, brand, p.category, p.asin, 'شراء', 'سعر', 'أونلاين', 'مصر'].filter(Boolean).join(', ');
  const priceStr = p.price ? p.price.toFixed(2) : '';
  const availability = /تبقى|محدودة|نفد/i.test(p.availability) ? 'LimitedAvailability' : 'InStock';
  const schemaAvailability = availability === 'LimitedAvailability' ? 'https://schema.org/LimitedAvailability' : 'https://schema.org/InStock';

  const jsonLd = {
    "@context": "https://schema.org/",
    "@type": "Product",
    "name": title,
    "image": p.images.map(hiRes).filter(Boolean),
    "description": descShort,
    "sku": p.asin,
    "mpn": p.asin,
    "brand": { "@type": "Brand", "name": brand },
    "category": p.category,
    ...(p.rating ? {
      "aggregateRating": {
        "@type": "AggregateRating",
        "ratingValue": String(p.rating),
        "reviewCount": String(p.reviews || 1),
        "bestRating": "5"
      }
    } : {}),
    ...(p.price ? {
      "offers": {
        "@type": "Offer",
        "url": pageUrl,
        "priceCurrency": CONFIG.currency,
        "price": priceStr,
        "availability": schemaAvailability,
        "itemCondition": "https://schema.org/NewCondition",
        "seller": { "@type": "Organization", "name": CONFIG.siteName }
      }
    } : {})
  };

  const breadcrumbLd = {
    "@context": "https://schema.org",
    "@type": "BreadcrumbList",
    "itemListElement": [
      { "@type": "ListItem", "position": 1, "name": "الرئيسية", "item": CONFIG.mainSiteUrl },
      { "@type": "ListItem", "position": 2, "name": p.category, "item": CONFIG.mainSiteUrl + '#cat-' + encodeURIComponent(p.category) },
      { "@type": "ListItem", "position": 3, "name": title, "item": pageUrl }
    ]
  };

  const related = PRODUCTS_CACHE
    .filter(x => x.category === p.category && x.asin !== p.asin)
    .slice(0, 4);

  return `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>${escapeHTML(title)} | ${escapeHTML(brand)} — ${escapeHTML(CONFIG.siteName)}</title>
<meta name="description" content="${escapeAttr(descShort)}">
<meta name="keywords" content="${escapeAttr(keywords)}">
<meta name="robots" content="index, follow, max-image-preview:large">
<meta name="author" content="${escapeAttr(brand)}">
<link rel="canonical" href="${escapeAttr(pageUrl)}">

<!-- Open Graph -->
<meta property="og:type" content="product">
<meta property="og:site_name" content="${escapeAttr(CONFIG.siteName)}">
<meta property="og:title" content="${escapeAttr(title)}">
<meta property="og:description" content="${escapeAttr(descShort)}">
<meta property="og:url" content="${escapeAttr(pageUrl)}">
<meta property="og:image" content="${escapeAttr(img)}">
<meta property="og:image:alt" content="${escapeAttr(title)}">
<meta property="og:locale" content="${CONFIG.locale}">
${p.price ? `<meta property="product:price:amount" content="${priceStr}">
<meta property="product:price:currency" content="${CONFIG.currency}">` : ''}

<!-- Twitter -->
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:title" content="${escapeAttr(title)}">
<meta name="twitter:description" content="${escapeAttr(descShort)}">
<meta name="twitter:image" content="${escapeAttr(img)}">

<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>🛍️</text></svg>">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&display=swap" rel="stylesheet">

<script type="application/ld+json">${JSON.stringify(jsonLd)}</script>
<script type="application/ld+json">${JSON.stringify(breadcrumbLd)}</script>

<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Tajawal',system-ui,-apple-system,sans-serif;background:#f2f4f8;color:#0f172a;line-height:1.8}
a{text-decoration:none;color:inherit}
img{max-width:100%;display:block}
.wrap{max-width:1200px;margin:0 auto;padding:22px}
.top{background:#0b1220;color:#fff;padding:14px 22px;display:flex;align-items:center;gap:14px;box-shadow:0 4px 20px rgba(0,0,0,.15)}
.top .logo{font-size:22px;font-weight:900;display:flex;align-items:center;gap:9px}
.top .logo .mark{width:34px;height:34px;border-radius:11px;background:linear-gradient(135deg,#ff6a00,#ff9500);display:grid;place-items:center;font-size:18px}
.top a.back{margin-inline-start:auto;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:#fff;padding:8px 18px;border-radius:10px;font-weight:700;font-size:14px}
.top a.back:hover{background:#ff6a00;border-color:#ff6a00}
.crumb{font-size:13.5px;color:#6b7280;margin-bottom:18px;padding:8px 0}
.crumb a:hover{color:#ff6a00}
.crumb .sep{margin:0 8px;color:#cbd5e1}
.main{background:#fff;border-radius:22px;padding:30px;display:grid;grid-template-columns:1fr 1fr;gap:36px;box-shadow:0 8px 30px rgba(16,24,40,.08);margin-bottom:22px}
.gallery-main{background:linear-gradient(160deg,#fbfcfe,#eef2f8);border-radius:16px;aspect-ratio:1/1;display:grid;place-items:center;padding:22px;position:relative;overflow:hidden}
.gallery-main img{width:100%;height:100%;object-fit:contain;mix-blend-mode:multiply}
.gallery-thumbs{display:flex;gap:10px;margin-top:14px;overflow-x:auto;padding-bottom:4px}
.gallery-thumbs img{width:74px;height:74px;object-fit:contain;background:#fafbfd;border:2px solid #e8ebf0;border-radius:12px;padding:6px;cursor:pointer;mix-blend-mode:multiply;transition:.2s}
.gallery-thumbs img:hover{border-color:#ffc38a}
.info h1{font-size:26px;font-weight:800;line-height:1.5;margin-bottom:12px;color:#0f172a}
.brand-pill{display:inline-block;background:#fff4e9;color:#c2410c;padding:5px 14px;border-radius:100px;font-size:13px;font-weight:800;margin-bottom:12px;border:1px solid #ffe1c4}
.rate{display:flex;align-items:center;gap:10px;font-size:14px;color:#6b7280;margin-bottom:14px;flex-wrap:wrap}
.stars{color:#ffb400;font-size:16px;letter-spacing:2px;font-family:Arial}
.price-box{background:linear-gradient(150deg,#fff7f0,#fff);border:1px solid #ffe1c4;border-radius:16px;padding:18px 20px;margin-bottom:16px}
.price{font-size:36px;font-weight:900;color:#d81b3c;line-height:1;display:inline-block}
.price small{font-size:16px;font-weight:700;margin-inline-start:4px}
.old{font-size:16px;color:#9aa3b5;text-decoration:line-through;margin-inline-start:14px;font-weight:600}
.off-badge{background:linear-gradient(135deg,#e11d48,#ff4d4d);color:#fff;font-weight:900;font-size:13px;padding:5px 12px;border-radius:8px;margin-inline-start:10px;vertical-align:middle}
.tax{display:block;font-size:12.5px;color:#6b7280;margin-top:6px}
.hl{background:#f7faff;border:1px solid #e3edff;border-radius:14px;padding:14px 16px;margin-bottom:14px}
.hl h3{font-size:15px;font-weight:900;color:#1e3a8a;margin-bottom:9px}
.hl ul{list-style:none;display:flex;flex-direction:column;gap:6px}
.hl li{font-size:14px;color:#334155;display:flex;gap:8px}
.hl li::before{content:"✓";flex-shrink:0;width:18px;height:18px;border-radius:50%;background:#dbeafe;color:#1d4ed8;font-size:11px;font-weight:900;display:grid;place-items:center;margin-top:4px}
.meta{display:flex;flex-direction:column;gap:8px;font-size:14px;margin-bottom:16px;color:#374151}
.meta b{color:#0f172a;font-weight:800}
.buy{display:flex;gap:10px;flex-wrap:wrap}
.btn-buy{flex:1;min-width:180px;background:linear-gradient(135deg,#ff6a00,#ff9500);color:#fff;padding:15px 26px;border-radius:13px;font-weight:900;font-size:16px;text-align:center;box-shadow:0 12px 26px -10px rgba(255,106,0,.85);transition:.2s}
.btn-buy:hover{transform:translateY(-3px)}
.btn-store{padding:14px 24px;border-radius:13px;font-weight:900;font-size:15px;background:#0f172a;color:#fff;border:2px solid #0f172a;transition:.2s}
.btn-store:hover{background:#1e293b;transform:translateY(-3px)}
.tabs{margin-top:8px}
.tab-panel{background:#fff;border-radius:22px;padding:26px 30px;box-shadow:0 8px 30px rgba(16,24,40,.08);margin-bottom:22px}
.tab-panel h2{font-size:20px;font-weight:900;margin-bottom:16px;color:#0f172a;display:flex;align-items:center;gap:10px}
.tab-panel h2::before{content:"";width:5px;height:22px;background:linear-gradient(180deg,#ff6a00,#ff9500);border-radius:3px}
.desc{font-size:16px;line-height:2;color:#374151;white-space:pre-line}
.specs{width:100%;border-collapse:collapse;font-size:14.5px}
.specs tr{border-bottom:1px solid #e8ebf0}
.specs tr:nth-child(odd){background:#fafbfd}
.specs td{padding:12px 16px;vertical-align:top}
.specs td:first-child{width:34%;font-weight:800;color:#475569;background:#f6f8fb;border-inline-end:1px solid #e8ebf0}
.specs td:last-child{color:#1f2937;font-weight:600}
.related{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:16px}
.rel-card{background:#fff;border:1px solid #e8ebf0;border-radius:14px;overflow:hidden;transition:.2s;display:block}
.rel-card:hover{transform:translateY(-4px);box-shadow:0 12px 26px rgba(16,24,40,.1);border-color:#ffd9b8}
.rel-card img{aspect-ratio:1/1;object-fit:contain;background:#fafbfd;padding:12px;mix-blend-mode:multiply}
.rel-card .body{padding:11px 13px 14px}
.rel-card .t{font-size:13.5px;font-weight:700;line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;color:#1f2937;min-height:40px}
.rel-card .p{margin-top:8px;font-size:17px;font-weight:900;color:#d81b3c}
.rel-card .p small{font-size:12px;font-weight:700;margin-inline-start:3px}
footer{background:#0b1220;color:#94a3b8;padding:30px 22px;text-align:center;font-size:13.5px;margin-top:22px}
footer a{color:#ff9500;font-weight:700}
@media(max-width:860px){
  .main{grid-template-columns:1fr;padding:20px;gap:22px}
  .info h1{font-size:20px}
  .price{font-size:28px}
  .wrap{padding:14px}
  .tab-panel{padding:20px 18px}
}
</style>
</head>
<body>

<header class="top">
  <a href="${escapeAttr(CONFIG.mainSiteUrl)}" class="logo">
    <span class="mark">🛍️</span> ${escapeHTML(CONFIG.siteName)}
  </a>
  <a href="${escapeAttr(CONFIG.mainSiteUrl)}" class="back">← العودة للمتجر</a>
</header>

<div class="wrap">
  <div class="crumb">
    <a href="${escapeAttr(CONFIG.mainSiteUrl)}">الرئيسية</a>
    <span class="sep">›</span>
    <a href="${escapeAttr(CONFIG.mainSiteUrl)}">${escapeHTML(p.category)}</a>
    <span class="sep">›</span>
    <span>${escapeHTML(title.slice(0,60))}${title.length > 60 ? '...' : ''}</span>
  </div>

  <article class="main">
    <div class="gallery">
      <div class="gallery-main">
        <img id="mainImg" src="${escapeAttr(img)}" alt="${escapeAttr(title)}" loading="eager">
      </div>
      ${p.images.length > 1 ? `<div class="gallery-thumbs">${p.images.map((u,i)=>`<img src="${escapeAttr(hiRes(u))}" data-full="${escapeAttr(hiRes(u))}" alt="${escapeAttr(title)} - صورة ${i+1}" loading="lazy"${i===0?' style="border-color:#ff6a00"':''}>`).join('')}</div>` : ''}
    </div>

    <div class="info">
      <span class="brand-pill">🏷️ ${escapeHTML(brand)}</span>
      <h1>${escapeHTML(title)}</h1>

      ${p.rating ? `<div class="rate"><span class="stars">${'★'.repeat(Math.round(p.rating))}${'☆'.repeat(5-Math.round(p.rating))}</span><b style="color:#0f172a">${p.rating}</b> من 5 ${p.reviews ? `— ${p.reviews.toLocaleString('ar-EG')} تقييم` : ''}</div>` : ''}

      ${p.price ? `
      <div class="price-box">
        <span class="price">${p.price.toFixed(2)}<small>${CONFIG.currencyAr}</small></span>
        ${p.oldPrice > p.price ? `<span class="old">${p.oldPrice.toFixed(2)}</span>` : ''}
        ${p.discount ? `<span class="off-badge">وفّر ${p.discount}%</span>` : ''}
        <span class="tax">الأسعار تشمل ضريبة القيمة المضافة</span>
      </div>` : ''}

      ${p.highlights.length ? `<div class="hl"><h3>✨ أبرز المميزات</h3><ul>${p.highlights.map(h=>`<li>${escapeHTML(h)}</li>`).join('')}</ul></div>` : ''}

      <div class="meta">
        <div>📦 الحالة: <b>${escapeHTML(p.availability)}</b></div>
        ${p.delivery ? `<div>🚚 ${escapeHTML(p.delivery)}</div>` : ''}
        ${p.returns ? `<div>↩️ ${escapeHTML(p.returns)}</div>` : ''}
        ${p.seller ? `<div>🏪 البائع: <b>${escapeHTML(p.seller)}</b></div>` : ''}
        <div>🔖 ASIN: <b>${escapeHTML(p.asin)}</b></div>
      </div>

      <div class="buy">
        ${p.link && p.link !== '#' ? `<a class="btn-buy" href="${escapeAttr(p.link)}" target="_blank" rel="noopener nofollow sponsored">🛒 اشترِ الآن من المتجر</a>` : ''}
        <a class="btn-store" href="${escapeAttr(CONFIG.mainSiteUrl)}">تصفح منتجات أخرى</a>
      </div>
    </div>
  </article>

  ${p.description ? `
  <section class="tab-panel">
    <h2>📝 وصف المنتج</h2>
    <div class="desc">${escapeHTML(p.description)}</div>
  </section>` : ''}

  ${p.specs.length ? `
  <section class="tab-panel">
    <h2>📋 المواصفات الكاملة</h2>
    <table class="specs"><tbody>
      ${p.specs.map(([k,v])=>`<tr><td>${escapeHTML(k)}</td><td>${escapeHTML(v)}</td></tr>`).join('')}
    </tbody></table>
  </section>` : ''}

  ${related.length ? `
  <section class="tab-panel">
    <h2>🛍️ منتجات مشابهة</h2>
    <div class="related">
      ${related.map(r=>`
        <a class="rel-card" href="${escapeAttr(r.asin)}.html">
          <img src="${escapeAttr(hiRes(r.image))}" alt="${escapeAttr(r.title)}" loading="lazy">
          <div class="body">
            <div class="t">${escapeHTML(r.title)}</div>
            ${r.price ? `<div class="p">${r.price.toFixed(2)}<small>${CONFIG.currencyAr}</small></div>` : ''}
          </div>
        </a>`).join('')}
    </div>
  </section>` : ''}
</div>

<footer>
  © ${new Date().getFullYear()} <a href="${escapeAttr(CONFIG.mainSiteUrl)}">${escapeHTML(CONFIG.siteName)}</a> — ${escapeHTML(CONFIG.siteTagline)}
</footer>

<script>
document.querySelectorAll('.gallery-thumbs img').forEach(t => {
  t.addEventListener('click', () => {
    document.getElementById('mainImg').src = t.dataset.full;
    document.querySelectorAll('.gallery-thumbs img').forEach(x => x.style.borderColor = '#e8ebf0');
    t.style.borderColor = '#ff6a00';
  });
});
</script>
</body>
</html>`;
}

/* ============ توليد index للـ SEO (صفحة رئيسية ثابتة) ============ */
function generateSEOIndex(products){
  const cats = {};
  products.forEach(p => { (cats[p.category] = cats[p.category] || []).push(p); });

  return `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>كل المنتجات | ${escapeHTML(CONFIG.siteName)}</title>
<meta name="description" content="${escapeAttr(CONFIG.siteTagline)} — تصفح ${products.length}+ منتج مع الأسعار والمواصفات الكاملة.">
<meta name="robots" content="index, follow">
<link rel="canonical" href="${escapeAttr(CONFIG.siteUrl + '/product/index.html')}">
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700;800;900&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Tajawal',system-ui,sans-serif;background:#f2f4f8;color:#0f172a;line-height:1.7}
a{text-decoration:none;color:inherit}
.top{background:#0b1220;color:#fff;padding:16px 22px;text-align:center}
.top h1{font-size:22px;font-weight:900}
.top p{font-size:14px;color:#94a3b8;margin-top:4px}
.wrap{max-width:1400px;margin:22px auto;padding:0 22px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:18px}
.card{background:#fff;border:1px solid #e8ebf0;border-radius:16px;overflow:hidden;transition:.2s;display:block}
.card:hover{transform:translateY(-4px);box-shadow:0 12px 30px rgba(16,24,40,.12);border-color:#ffd9b8}
.card img{aspect-ratio:1/1;object-fit:contain;background:#fafbfd;padding:14px;mix-blend-mode:multiply;width:100%}
.body{padding:13px 14px 16px}
.brand{font-size:12px;color:#ff6a00;font-weight:800;margin-bottom:4px}
.t{font-size:14px;font-weight:700;line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;color:#1f2937;min-height:42px}
.p{margin-top:8px;font-size:19px;font-weight:900;color:#d81b3c}
.p small{font-size:12px;font-weight:700;margin-inline-start:3px}
</style>
</head>
<body>
<div class="top"><h1>🛍️ ${escapeHTML(CONFIG.siteName)} — كل المنتجات</h1><p>${products.length} منتج متوفر</p></div>
<div class="wrap">
  <div class="grid">
    ${products.map(p => `
      <a class="card" href="${escapeAttr(p.asin)}.html">
        <img src="${escapeAttr(hiRes(p.image))}" alt="${escapeAttr(p.title)}" loading="lazy">
        <div class="body">
          <div class="brand">${escapeHTML(p.brand || '')}</div>
          <div class="t">${escapeHTML(p.title)}</div>
          ${p.price ? `<div class="p">${p.price.toFixed(2)}<small>${CONFIG.currencyAr}</small></div>` : ''}
        </div>
      </a>`).join('')}
  </div>
</div>
</body>
</html>`;
}

/* ============ توليد sitemap.xml ============ */
function generateSitemap(products){
  const base = CONFIG.siteUrl.replace(/\/$/,'');
  const today = new Date().toISOString().split('T')[0];
  let xml = `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n`;
  xml += `  <url>\n    <loc>${escapeHTML(base + '/' + CONFIG.outputDir + '/index.html')}</loc>\n    <lastmod>${today}</lastmod>\n    <changefreq>daily</changefreq>\n    <priority>0.9</priority>\n  </url>\n`;
  for(const p of products){
    xml += `  <url>\n    <loc>${escapeHTML(base + '/' + CONFIG.outputDir + '/' + p.asin + '.html')}</loc>\n    <lastmod>${today}</lastmod>\n    <changefreq>weekly</changefreq>\n    <priority>0.8</priority>\n  </url>\n`;
  }
  xml += `</urlset>\n`;
  return xml;
}

/* ============ توليد robots.txt ============ */
function generateRobots(){
  const base = CONFIG.siteUrl.replace(/\/$/,'');
  return `User-agent: *\nAllow: /\n\nSitemap: ${base}/sitemap.xml\n`;
}

/* ============ حلقة التوليد ============ */
let PRODUCTS_CACHE = [];

function runOnce(){
  console.log('\n🔍 قراءة الملف:', CONFIG.excelFile);
  const products = readExcel();
  PRODUCTS_CACHE = products;
  console.log(`✅ تم العثور على ${products.length} منتج`);

  // إنشاء مجلد الإخراج
  if(!fs.existsSync(CONFIG.outputDir)) fs.mkdirSync(CONFIG.outputDir, { recursive:true });

  // توليد صفحة لكل منتج
  let generated = 0;
  for(const p of products){
    const file = path.join(CONFIG.outputDir, p.asin + '.html');
    fs.writeFileSync(file, generateProductPage(p), 'utf8');
    generated++;
  }
  console.log(`📄 تم توليد ${generated} صفحة منتج في ./${CONFIG.outputDir}/`);

  // صفحة index ثابتة
  fs.writeFileSync(path.join(CONFIG.outputDir, 'index.html'), generateSEOIndex(products), 'utf8');
  console.log(`📄 تم توليد صفحة الفهرس ./${CONFIG.outputDir}/index.html`);

  // sitemap.xml
  fs.writeFileSync('sitemap.xml', generateSitemap(products), 'utf8');
  console.log(`🗺️  تم توليد sitemap.xml`);

  // robots.txt
  fs.writeFileSync('robots.txt', generateRobots(), 'utf8');
  console.log(`🤖 تم توليد robots.txt`);

  console.log(`\n✨ اكتمل التوليد بنجاح — ${new Date().toLocaleTimeString('ar-EG')}\n`);
}

/* watch mode */
function runWatch(){
  let lastMtime = 0;
  console.log(`👀 وضع المراقبة مُفعّل — سيتم فحص "${CONFIG.excelFile}" كل ${CONFIG.watchIntervalMs/1000} ثانية`);
  runOnce();
  try{ lastMtime = fs.statSync(CONFIG.excelFile).mtimeMs; }catch(e){}
  setInterval(() => {
    try{
      const mtime = fs.statSync(CONFIG.excelFile).mtimeMs;
      if(mtime !== lastMtime){
        lastMtime = mtime;
        console.log(`\n📢 تم اكتشاف تعديل على الملف — إعادة التوليد...`);
        runOnce();
      }
    }catch(e){ console.error('خطأ في المراقبة:', e.message); }
  }, CONFIG.watchIntervalMs);
}

/* تشغيل */
if(CONFIG.watchMode) runWatch();
else runOnce();