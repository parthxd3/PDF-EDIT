// Prints four realistic HTML documents to PDF with headless Chrome (Chrome print-to-PDF producer).
const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');
const OUT = process.argv[2];
const PICS = process.argv[3];
const logo = 'data:image/png;base64,' + fs.readFileSync(path.join(PICS, 'logo.png')).toString('base64');

const base = `*{box-sizing:border-box}body{margin:0;color:#222;-webkit-print-color-adjust:exact;print-color-adjust:exact}`;

const docs = {};

docs.invoice_web = `<!doctype html><html><head><meta charset="utf-8"><style>${base}
@page{size:A4;margin:0}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:10pt}
.band{background:#1f3a5f;color:#fff;padding:28px 44px 24px;display:flex;justify-content:space-between;align-items:center}
.brand{display:flex;align-items:center;gap:14px}.brand img{width:54px;height:54px}
.brand h1{margin:0;font-size:20pt;font-weight:600}.brand div{font-size:9pt;opacity:.85}
.band .ttl{text-align:right}.band .ttl h2{margin:0;font-size:26pt;letter-spacing:3px;font-weight:300}
.band .ttl span{font-size:9pt;opacity:.85}
.wrap{padding:28px 44px}
.meta{display:flex;gap:24px;margin-bottom:26px}
.meta .box{flex:1;border:1px solid #d9dee5;border-radius:6px;padding:12px 14px}
.meta h3{margin:0 0 6px;font-size:8pt;text-transform:uppercase;letter-spacing:1px;color:#5b6b80}
.meta p{margin:0;line-height:1.45}
table{width:100%;border-collapse:collapse}
th{background:#eef2f7;color:#1f3a5f;text-align:left;font-size:8.5pt;text-transform:uppercase;letter-spacing:.5px;padding:9px 10px;border-bottom:2px solid #1f3a5f}
td{padding:9px 10px;border-bottom:1px solid #e3e7ec;vertical-align:top}
td.n,th.n{text-align:right;white-space:nowrap}
td small{color:#6b7785;display:block;font-size:8.5pt}
.sum{display:flex;justify-content:space-between;margin-top:22px;gap:30px}
.note{flex:1;font-size:9pt;color:#4a5562;line-height:1.5}
.note h4{margin:0 0 4px;color:#1f3a5f;font-size:9.5pt}
.totals{width:250px;border:1px solid #d9dee5;border-radius:6px;overflow:hidden}
.totals div{display:flex;justify-content:space-between;padding:8px 14px;border-bottom:1px solid #e3e7ec}
.totals .grand{background:#1f3a5f;color:#fff;font-weight:700;font-size:12pt;border:0}
.foot{position:absolute;bottom:26px;left:44px;right:44px;border-top:1px solid #d9dee5;padding-top:8px;font-size:8pt;color:#7a8592;display:flex;justify-content:space-between}
</style></head><body>
<div class="band"><div class="brand"><img src="${logo}"><div><h1>Halden Cloud Services</h1><div>Managed hosting and IT support</div></div></div>
<div class="ttl"><h2>INVOICE</h2><span>No. HCS-24117 &nbsp;|&nbsp; 7 November 2026</span></div></div>
<div class="wrap">
<div class="meta">
<div class="box"><h3>Billed to</h3><p><b>Ardent Physiotherapy Ltd</b><br>Finance Department<br>Suite 3, 41 Castle Street<br>Shrewsbury SY1 2BW</p></div>
<div class="box"><h3>From</h3><p><b>Halden Cloud Services Ltd</b><br>12 Wharf Road<br>Birmingham B1 2DS<br>VAT GB 274 9918 32</p></div>
<div class="box"><h3>Details</h3><p>Account: ARD-0093<br>Period: 1 to 31 October 2026<br>Terms: 14 days<br><b>Due: 21 November 2026</b></p></div>
</div>
<table><thead><tr><th>Description</th><th class="n">Qty</th><th class="n">Rate</th><th class="n">Amount</th></tr></thead><tbody>
<tr><td>Managed server, standard tier<small>2 vCPU, 8 GB RAM, daily backups, 99.9% uptime</small></td><td class="n">2</td><td class="n">£85.00</td><td class="n">£170.00</td></tr>
<tr><td>Microsoft 365 Business Standard<small>Monthly licence, per user</small></td><td class="n">14</td><td class="n">£10.30</td><td class="n">£144.20</td></tr>
<tr><td>Helpdesk support plan<small>Business hours, 4-hour response</small></td><td class="n">1</td><td class="n">£220.00</td><td class="n">£220.00</td></tr>
<tr><td>On-site visit: network switch replacement<small>Engineer time, 3.5 hours on 16 October</small></td><td class="n">3.5</td><td class="n">£65.00</td><td class="n">£227.50</td></tr>
<tr><td>Cisco CBS250 24-port switch<small>Hardware, supplied and installed</small></td><td class="n">1</td><td class="n">£312.00</td><td class="n">£312.00</td></tr>
<tr><td>Domain renewal: ardentphysio.example<small>12 months</small></td><td class="n">1</td><td class="n">£14.99</td><td class="n">£14.99</td></tr>
</tbody></table>
<div class="sum">
<div class="note"><h4>How to pay</h4>Bank transfer to Halden Cloud Services Ltd, sort code 30-91-56, account 17730442. Please use the invoice number as your reference. Card payments can be made through the customer portal.<br><br><h4>Questions?</h4>Contact accounts@haldencloud.example or call 0121 496 0388.</div>
<div class="totals"><div><span>Subtotal</span><span>£1,088.69</span></div><div><span>VAT (20%)</span><span>£217.74</span></div><div><span>Paid</span><span>£0.00</span></div><div class="grand"><span>Amount due</span><span>£1,306.43</span></div></div>
</div></div>
<div class="foot"><span>Halden Cloud Services Ltd, registered in England no. 10288417</span><span>Page 1 of 1</span></div>
</body></html>`;

docs.brochure = `<!doctype html><html><head><meta charset="utf-8"><style>${base}
@page{size:A4;margin:16mm 0 14mm}@page :first{margin-top:0}
body{font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;line-height:1.5}
.hero{background:linear-gradient(135deg,#0e5e6f,#3a8891);color:#fff;padding:54px 50px 46px}
.hero .kicker{text-transform:uppercase;letter-spacing:3px;font-size:9pt;opacity:.85}
.hero h1{font-family:Georgia,serif;font-size:32pt;line-height:1.15;margin:10px 0 14px;font-weight:normal}
.hero p{font-size:12pt;max-width:470px;margin:0 0 22px}
.btn{display:inline-block;background:#f2b134;color:#1d2b2f;font-weight:bold;padding:10px 22px;border-radius:22px;font-size:10pt}
.sec{padding:34px 50px 10px}
.sec h2{font-family:Georgia,serif;font-size:18pt;color:#0e5e6f;margin:0 0 6px;font-weight:normal}
.sec .lead{color:#555;margin:0 0 20px}
.cards{display:flex;gap:18px}
.card{flex:1;border:1px solid #dfe7e8;border-top:5px solid #0e5e6f;border-radius:6px;padding:18px 16px;background:#f7fafa}
.card .ic{width:34px;height:34px;border-radius:50%;background:#0e5e6f;color:#fff;text-align:center;line-height:34px;font-weight:bold;margin-bottom:8px}
.card h3{margin:0 0 6px;font-size:12pt;color:#1d2b2f}
.card p{margin:0;font-size:9.5pt;color:#444}
.cols{column-count:2;column-gap:30px;text-align:justify}
.cols p{margin:0 0 10px}
.cols h3{font-size:11.5pt;color:#0e5e6f;margin:0 0 4px;break-after:avoid}
.stats{display:flex;background:#1d2b2f;color:#fff;margin:24px 50px 0;border-radius:8px}
.stats div{flex:1;text-align:center;padding:16px 8px}
.stats b{display:block;font-size:20pt;color:#f2b134;font-family:Georgia,serif}
.stats span{font-size:8.5pt;opacity:.85}
.contact{margin:26px 50px;padding-top:14px;border-top:2px solid #0e5e6f;display:flex;justify-content:space-between;font-size:9pt;color:#555}
</style></head><body>
<div class="hero"><div class="kicker">Greenway Home Energy</div><h1>Lower bills, warmer rooms<br>and a smaller carbon footprint</h1>
<p>Independent advice and installation of heat pumps, solar panels and insulation for homes across Yorkshire and the North East.</p><span class="btn">Book a free home survey</span></div>
<div class="sec"><h2>What we do</h2><p class="lead">One team from the first survey to the final certificate, so you never have to coordinate trades yourself.</p>
<div class="cards">
<div class="card"><div class="ic">1</div><h3>Heat pumps</h3><p>Air-source systems sized room by room, with the Boiler Upgrade Scheme grant of £7,500 deducted from your quote.</p></div>
<div class="card"><div class="ic">2</div><h3>Solar and batteries</h3><p>Panels designed around your roof and usage, with optional battery storage to use more of what you generate.</p></div>
<div class="card"><div class="ic">3</div><h3>Insulation</h3><p>Loft, cavity wall and underfloor insulation, often the cheapest way to cut heat loss before anything else.</p></div>
</div></div>
<div class="stats"><div><b>2,400+</b><span>homes upgraded since 2015</span></div><div><b>38%</b><span>average saving on heating bills</span></div><div><b>4.9/5</b><span>from 610 customer reviews</span></div><div><b>10 yrs</b><span>workmanship guarantee</span></div></div>
<div class="sec"><h2>How it works</h2>
<div class="cols">
<h3>1. Home survey</h3><p>One of our assessors visits your home for about two hours. We measure every room, check the walls, loft and windows, and look at your recent energy bills. You receive a written report within five working days, with costs and expected savings for each option.</p>
<h3>2. Design and quote</h3><p>If you decide to go ahead, our engineers produce a detailed design and a fixed quote. There are no hidden extras: the quote includes scaffolding, electrical work, building notifications and removal of old equipment.</p>
<h3>3. Installation</h3><p>Most heat pump installations take three to five days and solar installations one to two. We protect floors and furniture, keep you informed each day and leave the site clean. You will have heating and hot water every night during the work.</p>
<h3>4. Aftercare</h3><p>We register your system for the relevant guarantees and grants, show you how to use the controls and call you after the first month to check everything is working. Annual servicing is available from £120.</p>
</div></div>
<div class="contact"><span><b>Greenway Home Energy Ltd</b>, 8 Kirkgate, Ripon HG4 1PA</span><span>01765 600 482 &nbsp;|&nbsp; hello@greenwayenergy.example</span></div>
</body></html>`;

const box = (label, value, w) => `<div class="f" style="flex:${w || 1}"><label>${label}</label><div class="v">${value}</div></div>`;
const cb = (on, label) => `<span class="cb"><i class="${on ? 'on' : ''}">${on ? '&#10003;' : ''}</i>${label}</span>`;
docs.form_web = `<!doctype html><html><head><meta charset="utf-8"><style>${base}
@page{size:A4;margin:16mm 15mm}
body{font-family:Arial,Helvetica,sans-serif;font-size:9.5pt}
.head{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:3px solid #2b3a67;padding-bottom:8px;margin-bottom:12px}
.head h1{margin:0;font-size:16pt;color:#2b3a67}.head div{font-size:8.5pt;color:#555;text-align:right}
.intro{font-size:8.5pt;color:#444;margin:0 0 12px}
section{border:1px solid #9aa3b5;margin-bottom:12px}
section h2{margin:0;background:#2b3a67;color:#fff;font-size:9.5pt;padding:5px 8px;text-transform:uppercase;letter-spacing:.5px}
.row{display:flex;border-top:1px solid #c5cbd8}.row:first-of-type{border-top:0}
.f{padding:4px 8px 6px;border-left:1px solid #c5cbd8;min-height:38px}.f:first-child{border-left:0}
.f label{display:block;font-size:7.5pt;color:#5a6378;text-transform:uppercase}
.v{font-family:'Courier New',monospace;font-size:10pt;color:#0b2a8a;margin-top:3px}
.opts{padding:6px 8px;display:flex;flex-wrap:wrap;gap:6px 22px}
.cb{display:inline-flex;align-items:center;gap:6px}
.cb i{display:inline-block;width:12px;height:12px;border:1.2px solid #333;font-style:normal;font-size:9pt;line-height:11px;text-align:center;color:#0b2a8a;font-weight:bold}
.q{padding:6px 8px;font-size:9pt}
.long{padding:6px 8px;min-height:60px;font-family:'Courier New',monospace;color:#0b2a8a;font-size:9.5pt;line-height:1.45}
.sign{display:flex;gap:0}.sign .f{min-height:52px}
.sig{font-family:'Segoe Script','Brush Script MT',cursive;font-size:15pt;color:#0b2a8a}
.office{border-style:dashed;background:#f3f4f7}
</style></head><body>
<div class="head"><h1>Application for a Residents' Parking Permit</h1><div>Westbury District Council<br>Parking Services &nbsp;|&nbsp; Form RP1 (rev. 2026)</div></div>
<p class="intro">Please complete all sections in BLOCK CAPITALS. Attach proof of address dated within the last three months and the vehicle registration certificate (V5C). Incomplete applications will be returned.</p>
<section><h2>1. Applicant details</h2>
<div class="row">${box('Title', 'MR', .5)}${box('First name(s)', 'THOMAS EDWARD', 2)}${box('Surname', 'GALLAGHER', 2)}</div>
<div class="row">${box('Address', '14 ORCHARD TERRACE, WESTBURY', 3)}${box('Postcode', 'WB3 7HN', 1)}</div>
<div class="row">${box('Daytime telephone', '01632 960 274', 1)}${box('Email address', 'T.GALLAGHER@POSTBOX.EXAMPLE', 2)}${box('Date of birth', '19 / 04 / 1983', 1)}</div>
</section>
<section><h2>2. Permit type</h2>
<div class="opts">${cb(true, 'Resident permit (£48 per year)')}${cb(false, 'Visitor permits (book of 20, £15)')}${cb(false, 'Carer permit (free)')}${cb(false, 'Business permit (£180 per year)')}</div>
<div class="row">${box('Permit zone', 'ZONE C - ORCHARD / MILL LANE', 2)}${box('Start date', '01 / 12 / 2026', 1)}${box('Duration', cb(true, '12 months') + ' &nbsp; ' + cb(false, '6 months'), 1.4)}</div>
</section>
<section><h2>3. Vehicle details</h2>
<div class="row">${box('Registration number', 'WB19 KXP', 1)}${box('Make', 'SKODA', 1)}${box('Model', 'OCTAVIA ESTATE', 1)}${box('Colour', 'GREY', 1)}</div>
<div class="opts">The vehicle is: ${cb(true, 'Registered to me at this address')}${cb(false, 'A company or lease car')}${cb(false, 'Registered elsewhere')}</div>
</section>
<section><h2>4. Proof of residence enclosed</h2>
<div class="opts">${cb(true, 'Council tax bill')}${cb(false, 'Tenancy agreement')}${cb(true, 'Utility bill')}${cb(false, 'Bank statement')}${cb(false, 'Other (please state)')}</div>
<div class="q">Is there off-street parking (a driveway or garage) at your property? &nbsp; ${cb(false, 'Yes')} &nbsp; ${cb(true, 'No')}</div>
</section>
<section><h2>5. Additional information</h2>
<div class="long">WE MOVED TO ORCHARD TERRACE IN SEPTEMBER. MY PREVIOUS PERMIT (ZONE A, NO. 22817) SHOULD BE CANCELLED FROM THE START DATE OF THIS ONE.</div>
</section>
<section><h2>6. Declaration</h2>
<div class="q">I declare that the information given is correct. I understand that a permit obtained with false information will be cancelled and that the council may check the details with the DVLA.</div>
<div class="row sign">${box('Signature', '<span class="sig">T. Gallagher</span>', 2)}${box('Date', '09 / 11 / 2026', 1)}</div>
</section>
<section class="office"><h2 style="background:#7b8296">For office use only</h2>
<div class="row">${box('Received', '', 1)}${box('Permit no.', '', 1)}${box('Checked by', '', 1)}${box('Payment ref.', '', 1)}</div>
</section>
</body></html>`;

docs.cv_web = `<!doctype html><html><head><meta charset="utf-8"><style>${base}
@page{size:A4;margin:0}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:9.8pt;line-height:1.45}
.page{display:flex;min-height:297mm}
aside{width:72mm;background:#243447;color:#e8edf2;padding:34px 22px}
aside .ph{width:96px;height:96px;border-radius:50%;background:#4f6d8f;color:#fff;font-size:30pt;text-align:center;line-height:96px;margin:0 auto 16px;font-weight:300}
aside h2{font-size:9pt;text-transform:uppercase;letter-spacing:2px;color:#8fb3d9;border-bottom:1px solid #3d5470;padding-bottom:4px;margin:22px 0 8px}
aside p{margin:0 0 4px}
.bar{height:5px;background:#3d5470;border-radius:3px;margin:2px 0 8px}.bar i{display:block;height:5px;border-radius:3px;background:#8fb3d9}
aside ul{margin:0;padding-left:16px}
main{flex:1;padding:36px 34px}
main h1{margin:0;font-size:26pt;color:#243447;font-weight:600;letter-spacing:.5px}
main .role{font-size:12.5pt;color:#4f6d8f;margin-bottom:14px}
main h2{font-size:11pt;text-transform:uppercase;letter-spacing:1.5px;color:#243447;margin:20px 0 8px;padding-bottom:3px;border-bottom:2px solid #8fb3d9}
.job{margin-bottom:12px}.job .t{display:flex;justify-content:space-between}.job b{font-size:10.5pt}
.job .d{color:#6b7785;font-size:9pt;white-space:nowrap}.job .c{color:#4f6d8f;font-style:italic}
.job ul{margin:4px 0 0;padding-left:18px}.job li{margin-bottom:2px}
.summary{text-align:justify}
</style></head><body><div class="page">
<aside><div class="ph">AL</div>
<h2>Contact</h2><p>Bristol, UK</p><p>07700 900 562</p><p>amara.lindqvist@mail.example</p><p>github.com/alindqvist</p>
<h2>Skills</h2>
<p>TypeScript and React</p><div class="bar"><i style="width:95%"></i></div>
<p>Node.js and GraphQL</p><div class="bar"><i style="width:85%"></i></div>
<p>Accessibility (WCAG 2.2)</p><div class="bar"><i style="width:90%"></i></div>
<p>Design systems</p><div class="bar"><i style="width:80%"></i></div>
<p>Testing: Playwright, Jest</p><div class="bar"><i style="width:75%"></i></div>
<h2>Education</h2><p><b>BSc Computer Science</b></p><p>University of Bath, 2014</p><p>First Class Honours</p>
<h2>Languages</h2><ul><li>English (native)</li><li>Swedish (fluent)</li><li>German (basic)</li></ul>
<h2>Interests</h2><p>Open-source maintainer of a small date-picker library; volunteer code club tutor; fell running.</p>
</aside>
<main><h1>Amara Lindqvist</h1><div class="role">Lead Front-End Engineer</div>
<h2>Summary</h2><p class="summary">Front-end engineer with eleven years of experience building fast, accessible web applications for public services and retail. I lead small teams, set up design systems that other teams actually adopt, and care about measurable improvements in performance and usability rather than rewrites for their own sake.</p>
<h2>Experience</h2>
<div class="job"><div class="t"><b>Lead Front-End Engineer</b><span class="d">2022 – present</span></div><div class="c">Kestrel Digital, Bristol</div>
<ul><li>Lead a team of six building the online booking service for a national network of 140 clinics.</li><li>Introduced a shared component library now used by four product teams; cut new-page build time by about 40%.</li><li>Brought the main booking journey to WCAG 2.2 AA; task completion for screen-reader users rose from 62% to 91%.</li></ul></div>
<div class="job"><div class="t"><b>Senior Front-End Developer</b><span class="d">2018 – 2022</span></div><div class="c">Harbourside Retail Group, Bristol</div>
<ul><li>Rebuilt the product listing pages in React with server rendering; largest contentful paint fell from 4.1 s to 1.6 s.</li><li>Ran A/B tests on checkout that increased mobile conversion by 9%.</li><li>Mentored five junior developers, two of whom are now senior engineers.</li></ul></div>
<div class="job"><div class="t"><b>Web Developer</b><span class="d">2014 – 2018</span></div><div class="c">Avon Valley Housing Association, Bath</div>
<ul><li>Built the tenant self-service portal for repairs and rent payments, used by 18,000 households.</li><li>Moved the public website to a static site generator, cutting hosting costs by 70%.</li></ul></div>
<h2>Selected talks</h2>
<p><b>Design systems that survive the reorg</b> – State of the Browser, London, 2025</p>
<p><b>Accessible date pickers are possible</b> – Async JS, Brighton, 2023</p>
<h2>References</h2><p>Available on request.</p>
</main></div></body></html>`;

(async () => {
  const browser = await puppeteer.launch({ executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe', headless: 'new' });
  try {
    for (const [name, html] of Object.entries(docs)) {
      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: 'load' });
      const p = path.join(OUT, name + '.pdf');
      await page.pdf({ path: p, format: 'A4', printBackground: true, preferCSSPageSize: true });
      fs.writeFileSync(path.join(OUT, name + '.html'), html);
      console.log('ok', name);
      await page.close();
    }
  } finally { await browser.close(); }
})().catch(e => { console.error(e); process.exit(1); });
