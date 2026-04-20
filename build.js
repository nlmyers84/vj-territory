#!/usr/bin/env node
/**
 * Build script: reads Full Accounts Data.xlsx + existing HTML RAW_DATA contacts,
 * merges them, applies scoring/status rules, and generates both HTML map files.
 */
const fs = require('fs');
const XLSX = require('xlsx');

// ── 1. Extract existing contacts from current HTML RAW_DATA ──
function extractExistingData(htmlPath) {
  const html = fs.readFileSync(htmlPath, 'utf8');
  const m = html.match(/const RAW_DATA\s*=\s*(\[[\s\S]*?\]);/);
  if (!m) { console.warn('No RAW_DATA found in', htmlPath); return []; }
  try { return JSON.parse(m[1]); } catch(e) { console.warn('Parse error', e.message); return []; }
}

// We need to extract from the PREVIOUS version in git, since current file was already overwritten
const { execSync } = require('child_process');
let existingData;
try {
  // Go back to the original commit before any rebuilds (200cd4a)
  const oldHTML = execSync('git show 200cd4a:index.html', { encoding: 'utf8', maxBuffer: 50*1024*1024 });
  const m2 = oldHTML.match(/const RAW_DATA\s*=\s*(\[[\s\S]*?\]);/);
  existingData = m2 ? JSON.parse(m2[1]) : [];
} catch(e) {
  console.warn('Could not read previous HTML from git, trying current file');
  existingData = extractExistingData('index.html');
}

// Build MULTIPLE lookups for contact matching (fallback strategy)
// 1. Exact: company + address
// 2. Fallback: company + city
// 3. Last resort: company name only (if unique)
const contactsByExactKey = {};    // company|address -> contacts
const contactsByNameCity = {};    // company|city -> contacts
const contactsByName = {};        // company -> [contacts arrays] (may have dupes)

existingData.forEach(d => {
  if (!d.contacts || d.contacts.length === 0) return;
  const name = d.company.toLowerCase().trim();
  const addr = (d.address || '').toLowerCase().trim();
  const city = (d.city || '').toLowerCase().trim();

  contactsByExactKey[name + '|' + addr] = d.contacts;
  contactsByNameCity[name + '|' + city] = d.contacts;
  if (!contactsByName[name]) contactsByName[name] = [];
  contactsByName[name].push(d.contacts);
});

function findExistingContacts(company, address, city) {
  const name = company.toLowerCase().trim();
  const addr = (address || '').toLowerCase().trim();
  const cty = (city || '').toLowerCase().trim();

  // 1. Exact match: company + address
  const exact = contactsByExactKey[name + '|' + addr];
  if (exact) return exact;

  // 2. Company + city
  const byCity = contactsByNameCity[name + '|' + cty];
  if (byCity) return byCity;

  // 3. Company name only (if unique — only one location)
  const byName = contactsByName[name];
  if (byName && byName.length === 1) return byName[0];

  return null;
}

console.log(`Extracted ${Object.keys(contactsByExactKey).length} accounts with contacts from previous HTML`);

// ── 2. Read Excel ──
const wb = XLSX.readFile('Full Accounts Data.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws);
console.log(`Read ${rows.length} rows from Excel`);

// ── 3. Excel date helper ──
function excelDateToJS(serial) {
  if (!serial || typeof serial !== 'number') return null;
  // Excel epoch is 1900-01-01, but off by 1 due to Lotus 1-2-3 bug
  return new Date((serial - 25569) * 86400 * 1000);
}

// ── 4. Filter & process ──
const EXCLUDE_NAMES = ['U S POSTAL SERVICE', 'USPS', 'BLUECREST INC'];
const VALID_STATES = ['KS', 'MO', 'KANSAS', 'MISSOURI'];

function normalizeState(s) {
  if (!s) return '';
  const u = s.trim().toUpperCase();
  if (u === 'KANSAS' || u === 'KS') return 'KS';
  if (u === 'MISSOURI' || u === 'MO') return 'MO';
  return u;
}

// Parse YOY trend from the HTML img tag
function parseYOYTrend(val) {
  if (!val || typeof val !== 'string') return 'unknown';
  const v = val.toLowerCase();
  if (v.includes('rising') || v.includes('green')) return 'Rising';
  if (v.includes('declining') || v.includes('red')) return 'Declining';
  if (v.includes('#error')) return 'unknown';
  return 'unknown';
}

// ── 5. Opportunity-based scoring ──
// Grades reflect "how likely does this site need marking/coding equipment?"
// based primarily on WHAT THEY DO + HOW BIG THEY ARE — available for nearly
// every account. Install data is a bonus signal, not a requirement. A 200-person
// food plant with zero known installs should still be an A.
//
// CORE signals (~70% of score, available for most accounts):
//   Vertical / industry fit        — food/pharma need coders on every line
//   Plant size (employees)         — more people = more production lines
//   Company revenue banding        — bigger company = more volume
//
// BONUS signals (~30%, rewards having intel but absence doesn't punish):
//   Competitive printer installs   — confirmed need
//   VJ installed base              — expansion opportunity
//   Current VJ revenue             — wallet share growth
//   Upgrade Readiness              — near-term equipment opportunity

// Vertical fit: "If I walked into a typical site in this vertical,
// how many production/packaging lines would need a coder?"
function verticalFitScore(vertical, industry) {
  const v = (vertical || '').toLowerCase();
  const ind = (industry || '').toLowerCase();

  // HIGH: every line needs a coder, high volume, regulatory requirements
  if (v.match(/food|meat|poultry|dairy|egg|baked|cereal|fruit|vegetable|frozen|snack|candy|confect|pet food|animal feed/)) return 30;
  if (v.match(/beverage|brew/)) return 28;
  if (v.match(/pharma|medical|cosmetic|personal care/)) return 28;

  // GOOD: packaging-intensive, uses marking
  if (v.match(/packag/)) return 22;
  if (v.match(/chemical/)) return 20;
  if (v.match(/plastic|rubber|extrusion|wire|cable/)) return 18;
  if (v.match(/consumer|cpg/)) return 18;

  // MODERATE: some lines, less volume
  if (v.match(/auto|aero|transport|aircraft/)) return 15;
  if (v.match(/building|construct/)) return 14;
  if (v.match(/industrial|electric/)) return 13;

  // LOW: minimal marking needs
  if (v.match(/graphic|print|text|distribut|oem/)) return 8;

  // Unknown vertical — fall back to industry
  if (!v || v === 'empty' || v.match(/other|unknown/i)) {
    if (ind.match(/agri|food|pharma|chemical/)) return 20;
    if (ind.includes('manufact')) return 15;
    if (ind.match(/retail|transport|mineral/)) return 10;
    return 5;
  }

  // Catch-all (numeric SIC codes etc)
  if (ind.includes('manufact')) return 15;
  return 5;
}

function calcScore(r) {
  let pts = 0;

  // ── CORE: What they do + how big (drives ~70% of grade) ──

  // Vertical / industry fit
  pts += verticalFitScore(r['_correctedVertical'] || r['Vertical'], r['Industry']);

  // Plant size (employees)
  const employees = Number(r['Employees']) || 0;
  if (employees >= 500) pts += 25;
  else if (employees >= 200) pts += 22;
  else if (employees >= 100) pts += 18;
  else if (employees >= 50) pts += 14;
  else if (employees >= 20) pts += 9;
  else if (employees >= 5) pts += 5;
  else pts += 1;

  // Company revenue banding
  const b = r['Annual Revenue Banding - USD'] || '';
  if (b.includes('5B+') || b.includes('1B-5B')) pts += 15;
  else if (b.includes('500M-1B') || b.includes('250M-500M')) pts += 13;
  else if (b.includes('100M-250M') || b.includes('50M-100M')) pts += 11;
  else if (b.includes('25M-50M') || b.includes('10M-25M')) pts += 8;
  else if (b.includes('5M-10M')) pts += 6;
  else if (b.includes('1M-5M')) pts += 3;
  else pts += 1;

  // ── BONUS: Known intel (rewards data, doesn't punish absence) ──

  // Competitive installs — confirmed need
  const compTotal = Number(r['Total number of Competitive Printers']) || 0;
  if (compTotal > 10) pts += 10;
  else if (compTotal >= 5) pts += 7;
  else if (compTotal >= 1) pts += 4;

  // VJ installed base — expansion opportunity
  const vjPrinters = Number(r['Total Number of Active Printers(VJDB)']) || 0;
  if (vjPrinters >= 10) pts += 8;
  else if (vjPrinters >= 5) pts += 5;
  else if (vjPrinters >= 1) pts += 3;

  // Current VJ revenue
  const rev012 = Number(r['Rev (All) - 0-12 Month Rolling']) || 0;
  if (rev012 > 100000) pts += 6;
  else if (rev012 >= 50000) pts += 5;
  else if (rev012 >= 10000) pts += 3;
  else if (rev012 > 0) pts += 1;

  // Upgrade readiness
  if (r['Upgrade Readiness'] === 'Y') pts += 4;

  return pts;
}

// For old carried-forward accounts (no Excel data), score from what we have
function calcScoreOld(d) {
  let pts = 0;

  // Vertical fit (same logic, using old vertical names)
  const v = (d.vertical || '').toLowerCase();
  if (v.match(/food|bev|bak|pet|dairy|meat|fruit|vegetable|cereal|snack|confect/)) pts += 30;
  else if (v.match(/pharma|medical|biotech|animal health/)) pts += 28;
  else if (v.match(/packag/)) pts += 22;
  else if (v.match(/chemical/)) pts += 20;
  else if (v.match(/plastic|rubber/)) pts += 18;
  else if (v.match(/consumer|cpg/)) pts += 18;
  else if (v.match(/auto|aero|transport/)) pts += 15;
  else if (v.match(/building|construct/)) pts += 14;
  else if (v.match(/beverage|brew/)) pts += 28;
  else pts += 8; // unknown — give benefit of doubt

  // Old fit was curated — A-fit = someone decided this was a good target
  if (d.fit === 'A') pts += 12;
  else if (d.fit === 'B') pts += 6;

  // Revenue as proxy for existing relationship size
  const sales = d.sales || 0;
  if (sales > 100000) pts += 6;
  else if (sales >= 50000) pts += 5;
  else if (sales >= 10000) pts += 3;
  else if (sales > 0) pts += 1;

  // Having contacts = better researched, more likely a real opportunity
  if (d.contacts && d.contacts.length >= 3) pts += 3;
  else if (d.contacts && d.contacts.length >= 1) pts += 1;

  return pts;
}

// Grades assigned AFTER all accounts are scored, using percentile quartiles
// (see section 8c below)

// ── 6. Determine account status ──
// Gold = active (rev last 12mo)
// Red-orange = lapsed (Type=Active Customer but no rev last 12mo)
// Pink = dormant (has VJ equipment installed but no orders last 12mo, was never typed Active)
// Prospect = everything else
function getStatus(r) {
  const type = String(r['Type'] || '').trim();
  const rev012 = Number(r['Rev (All) - 0-12 Month Rolling']) || 0;
  const rev1224 = Number(r['Rev (All) - 12-24 Month Rolling']) || 0;
  const vjPrinters = Number(r['Total Number of Active Printers(VJDB)']) || 0;

  // Any account with revenue in the last 12 months = Active Customer (gold)
  if (rev012 > 0) return 'active';

  // Type says Active Customer but zero revenue last 12 months = Lapsed (red-orange)
  if (type === 'Active Customer') return 'lapsed';

  // Has VJ equipment but no recent orders = Dormant (pink)
  // These accounts bought systems from us but stopped ordering
  if (vjPrinters > 0) return 'dormant';

  // Had revenue 12-24 months ago = also dormant
  if (rev1224 > 0) return 'dormant';

  return 'prospect';
}

// ── 7. City-level coordinate averages for fallback ──
const cityCoords = {};
rows.forEach(r => {
  const lat = Number(r['Latitude (MA Shipping)']);
  const lng = Number(r['Longitude (MA Shipping)']);
  if (lat && lng && !isNaN(lat) && !isNaN(lng)) {
    const city = (r['Shipping City'] || '').trim().toLowerCase();
    const state = normalizeState(r['Shipping State/Province']);
    const key = city + '|' + state;
    if (!cityCoords[key]) cityCoords[key] = { sumLat: 0, sumLng: 0, count: 0 };
    cityCoords[key].sumLat += lat;
    cityCoords[key].sumLng += lng;
    cityCoords[key].count++;
  }
});

function getCityAvgCoords(city, state) {
  const key = (city || '').trim().toLowerCase() + '|' + state;
  const c = cityCoords[key];
  if (c && c.count > 0) return { lat: c.sumLat / c.count, lng: c.sumLng / c.count };
  return null;
}

// ── 8. Process all rows ──
const accounts = [];
const seenKeys = new Set();

rows.forEach(r => {
  const name = (r['Account Name'] || '').trim();
  if (!name) return;

  // Exclude USPS / Bluecrest
  const nameUpper = name.toUpperCase();
  if (EXCLUDE_NAMES.some(ex => nameUpper.includes(ex))) return;

  // Filter to KS/MO only
  const state = normalizeState(r['Shipping State/Province']);
  if (state !== 'KS' && state !== 'MO') return;

  // Coordinates
  let lat = Number(r['Latitude (MA Shipping)']) || 0;
  let lng = Number(r['Longitude (MA Shipping)']) || 0;

  if (!lat || !lng || isNaN(lat) || isNaN(lng)) {
    const fallback = getCityAvgCoords(r['Shipping City'], state);
    if (fallback) {
      lat = fallback.lat;
      lng = fallback.lng;
    } else {
      return; // skip accounts with no coords and no fallback
    }
  }

  // Dedup key: Account Name + Shipping Address
  const addr = (r['Shipping Address Line 1'] || '').trim();
  const dedupKey = (name + '|' + addr).toLowerCase();
  if (seenKeys.has(dedupKey)) return;
  seenKeys.add(dedupKey);

  // Fix misclassified verticals on the row BEFORE scoring
  const rawVert = (r['Vertical'] || '').trim().toLowerCase();
  if (rawVert.match(/^other|^unknown|^\d|^$/) || !rawVert) {
    const nameLC = name.toLowerCase();
    if (nameLC.match(/pet\b|petfood|animal feed|purina/)) r['_correctedVertical'] = 'Pet Food & Animal Feed';
    else if (nameLC.match(/food|bakery|baking|meat|packing co|cheese|candy|confect|snack|pizza|sausage|treats|jerky|gourmet/)) r['_correctedVertical'] = 'Other Food';
    else if (nameLC.match(/dairy|milk|cream|butter|yogurt/)) r['_correctedVertical'] = 'Dairy & Eggs';
    else if (nameLC.match(/beverage|brew|bottl|distill|wine|beer|soda|juice|coffee/)) r['_correctedVertical'] = 'Beverage';
    else if (nameLC.match(/pharma|drug|health|medical|biotech|vaccine/)) r['_correctedVertical'] = 'Pharma & Medical';
    else if (nameLC.match(/plastic|rubber|polymer/)) r['_correctedVertical'] = 'Plastics';
    else if (nameLC.match(/packag|container|carton|box|envelope|label/)) r['_correctedVertical'] = 'Packaging Materials';
    else if (nameLC.match(/auto|motor[^s]|vehicle|tire|brake|piston/)) r['_correctedVertical'] = 'Aero/Auto';
    else if (nameLC.match(/aero|aircraft|aviation|aerospace/)) r['_correctedVertical'] = 'Aero/Auto';
    else if (nameLC.match(/chemical|lubricant|paint|adhesive|coating/)) r['_correctedVertical'] = 'Chemicals';
  }

  const status = getStatus(r);
  const score = calcScore(r);
  const grade = '?'; // assigned later via percentile quartiles
  const rev012 = Number(r['Rev (All) - 0-12 Month Rolling']) || 0;
  const rev1224 = Number(r['Rev (All) - 12-24 Month Rolling']) || 0;
  const revYTD = Number(r['Rev (All) - YTD Rolling']) || 0;
  const revPrvYTD = Number(r['Rev (All) - Prv YTD Rolling']) || 0;
  const revYOYDiff = Number(r['Rev (All) - YOY Difference']) || 0;
  const trend = parseYOYTrend(r['Rev (All) - YOY Trend']);
  const upgradeReady = String(r['Upgrade Readiness'] || 'N').toUpperCase().trim() === 'Y';
  const isContractCustomer = r['Contract Customer'] === '1';
  const breakFixCount = Number(r['Break Fix WOLI Count']) || 0;
  const apRev012 = Math.round(Number(r['Rev (AP) - 0-12 Month Rolling']) || 0);
  const apRev1224 = Math.round(Number(r['Rev (AP) - 12-24 Month Rolling']) || 0);
  const compCIJ = Number(r['No of installed competitive CIJ Printers']) || 0;
  const compLaser = Number(r['No of installed competitive LASER']) || 0;
  let prodLines = Number(r['Total Number of Active Printers(VJDB)']) || Number(r['Total Number of Active Printers_CR']) || 0;
  if (prodLines === 99) prodLines = 0; // filter out placeholder/default values

  // VJ equipment breakdown by technology (use VJ ERP counts)
  const vjEquip = {};
  const cijTotal = (Number(r['CIJ-1000 LINE-VJ ERP'])||0) + (Number(r['CIJ-LEGACY-VJ ERP'])||0) + (Number(r['CIJ-SIMPLICITY-VJ ERP'])||0);
  if (cijTotal > 0) vjEquip.CIJ = cijTotal;
  const laserTotal = Number(r['LASER-VJ ERP']) || 0;
  if (laserTotal > 0) vjEquip.Laser = laserTotal;
  const lcmTotal = (Number(r['LCM-2300-VJ ERP'])||0) + (Number(r['LCM-OTHER-VJ ERP'])||0);
  if (lcmTotal > 0) vjEquip.LCM = lcmTotal;
  const ttoTotal = Number(r['TTO-VJ ERP']) || 0;
  if (ttoTotal > 0) vjEquip.TTO = ttoTotal;
  const tijTotal = Number(r['TIJ-VJ ERP']) || 0;
  if (tijTotal > 0) vjEquip.TIJ = tijTotal;
  const lpaTotal = Number(r['LPA-VJ ERP']) || 0;
  if (lpaTotal > 0) vjEquip.LPA = lpaTotal;
  const baTotal = Number(r['BINARY ARRAY-VJ ERP']) || 0;
  if (baTotal > 0) vjEquip['Binary Array'] = baTotal;

  // Per-technology last purchase window (parts + supplies revenue by tech)
  // Shows when they last ordered for each technology type
  const techActivity = {};
  function checkTechRev(tech, cols012, cols1224) {
    const rev012 = cols012.reduce((s, c) => s + (Number(r[c]) || 0), 0);
    const rev1224 = cols1224.reduce((s, c) => s + (Number(r[c]) || 0), 0);
    if (rev012 > 0) techActivity[tech] = { window: '0-12mo', rev: Math.round(rev012) };
    else if (rev1224 > 0) techActivity[tech] = { window: '12-24mo', rev: Math.round(rev1224) };
    else if (vjEquip[tech]) techActivity[tech] = { window: '24mo+', rev: 0 };
  }
  checkTechRev('CIJ',
    ['REV IB SUP CIJ 0 12 MON LC', 'REV_IB_PART_CIJ_0_12_MON_LC'],
    ['REV IB SUP CIJ 12 24 MON LC', 'REV_IB_PART_CIJ_12_24_MON_LC']);
  checkTechRev('Laser',
    ['REV_IB_SUP_LASER_0_12_MON_LC', 'REV_IB_PART_LASER_0_12_MON_LC'],
    ['REV_IB_SUP_LASER_12_24_MON_LC', 'REV_IB_PART_LASER_12_24_MON_LC']);
  checkTechRev('LCM',
    ['REV IB SUP LCM 0-12 MON LC', 'REV_IB_PART_LCM_0_12_MON_LC'],
    ['REV IB SUP LCM 12 24 MON LC', 'REV_IB_PART_LCM_12_24_MON_LC']);
  checkTechRev('TTO',
    ['REV_IB_SUP_TTO_0_12_MON_LC', 'REV_IB_PART_TTO_0_12_MON_LC'],
    ['REV_IB_SUP_TTO_12_24_MON_LC', 'REV_IB_PART_TTO_12_24_MON_LC']);
  checkTechRev('TIJ',
    ['REV_IB_SUP_TIJ_0_12_MON_LC', 'REV_IB_PART_TIJ_0_12_MON_LC__c'],
    ['REV_IB_SUP_TIJ_12_24_MON_LC', 'REV_IB_PART_TIJ_12_24_MON_LC']);
  checkTechRev('LPA',
    ['REV IB SUP LPA 0-12 MON LC', 'REV_IB_PART_LPA_0_12_MON_LC__c'],
    ['REV_IB_SUP_LPA_12_24_MON_LC', 'REV_IB_PART_LPA_12_24_MON_LC']);

  const city = (r['Shipping City'] || '').trim();
  const vertical = r['_correctedVertical'] || (r['Vertical'] || r['Industry'] || 'Other / Unknown').trim();

  // Contacts: merge existing + equipment contact from Excel
  let contacts = [];

  // Keep existing contacts (up to 5) — use fuzzy matching
  const existingContacts = findExistingContacts(name, addr, city);
  if (existingContacts) {
    contacts = existingContacts.slice(0, 10);
  }

  // Add Equipment Contact from Excel
  const eqName = (r['Equipment Contact'] || '').trim();
  if (eqName) {
    contacts.push({
      name: eqName,
      title: 'Equipment Contact',
      phone: (r['Equipment Contact Phone'] || '').trim(),
      email: (r['Equipment Contact Email'] || '').trim(),
      mobile: '',
      fit: '',
      dm: false
    });
  }

  // Opportunity notes
  const openOpps = Number(r['# of Open Opportunities']) || 0;
  const funnelVal = Number(r['Total Funnel Value (LC)']) || 0;
  let oppNote = '';
  if (openOpps > 0) oppNote = `${openOpps} open opportunit${openOpps === 1 ? 'y' : 'ies'}`;
  if (funnelVal > 0) oppNote += (oppNote ? ' | ' : '') + `Funnel: $${Math.round(funnelVal).toLocaleString()}`;

  accounts.push({
    company: name,
    address: addr,
    city,
    state,
    vertical,
    lat: Math.round(lat * 10000) / 10000,
    lng: Math.round(lng * 10000) / 10000,
    status, // 'active', 'lapsed', 'prospect'
    grade,
    score,
    rev012: Math.round(rev012),
    rev1224: Math.round(rev1224),
    revYTD: Math.round(revYTD),
    revPrvYTD: Math.round(revPrvYTD),
    revYOYDiff: Math.round(revYOYDiff),
    trend,
    upgradeReady,
    isContractCustomer,
    breakFixCount,
    apRev012,
    apRev1224,
    compCIJ,
    compLaser,
    prodLines,
    vjEquip,
    techActivity,
    oppNote,
    healthGrade: null, // computed below for active accounts
    healthScore: 0,
    contacts
  });
});

// ── 8b. Carry forward old KS/MO accounts, merging where possible ──

// Normalize a company name for fuzzy matching
function normName(name) {
  return name.toLowerCase()
    .replace(/['']/g, '')
    .replace(/\b(inc|llc|corp|co|company|corporation|ltd|lp|group|enterprises|industries|international|america|americas|usa|us|north america)\b\.?/g, '')
    .replace(/\s*[—–-]\s*.*(new|plant|facility|division|div|site|loc|bldg|building).*$/i, '')
    .replace(/\s*\(.*?\)\s*/g, ' ')
    .replace(/[^a-z0-9 ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Extract the "core" name — first 1-2 significant words (for brand-name matching)
// e.g. "Berry Global" and "Berry Plastics Corp" both -> "berry"
function coreName(name) {
  const norm = normName(name);
  const words = norm.split(' ').filter(w => w.length > 2);
  // Return first 1 word for short names, first 2 for longer
  return words.slice(0, words.length <= 3 ? 1 : 2).join(' ');
}

// Build multiple lookup indexes for matching
const accountsByNormNameCity = {};   // normName|city -> [indices]
const accountsByCoreNameCity = {};   // coreName|city -> [indices]
const accountsByNormName = {};       // normName -> [indices]

accounts.forEach((a, idx) => {
  const city = a.city.toLowerCase().trim();
  const nn = normName(a.company);
  const cn = coreName(a.company);

  const key1 = nn + '|' + city;
  if (!accountsByNormNameCity[key1]) accountsByNormNameCity[key1] = [];
  accountsByNormNameCity[key1].push(idx);

  const key2 = cn + '|' + city;
  if (!accountsByCoreNameCity[key2]) accountsByCoreNameCity[key2] = [];
  accountsByCoreNameCity[key2].push(idx);

  if (!accountsByNormName[nn]) accountsByNormName[nn] = [];
  accountsByNormName[nn].push(idx);
});

// Helper: merge old contacts into an existing account
function mergeContacts(existing, oldContacts) {
  const existingNames = new Set(existing.contacts.map(c => c.name.toLowerCase()));
  const newContacts = oldContacts.filter(c => !existingNames.has(c.name.toLowerCase()));
  // No cap during merge — keep all contacts so none are lost
  const totalSlots = 999;
  if (totalSlots > 0 && newContacts.length > 0) {
    const equipIdx = existing.contacts.findIndex(c => c.title === 'Equipment Contact');
    const toAdd = newContacts.slice(0, totalSlots);
    if (equipIdx >= 0) {
      existing.contacts.splice(equipIdx, 0, ...toAdd);
    } else {
      existing.contacts.push(...toAdd);
    }
    return true;
  }
  return false;
}

// Helper: normalize address for comparison
function normAddr(addr) {
  return (addr || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim();
}

// Helper: check if two addresses refer to the same physical site
function addrMatch(a, b) {
  const na = normAddr(a);
  const nb = normAddr(b);
  if (!na || !nb) return false; // can't confirm same site without address
  if (na === nb) return true;
  if (na.includes(nb) || nb.includes(na)) return true;
  // Check if street number matches (first numeric portion)
  const numA = na.match(/^(\d+)/);
  const numB = nb.match(/^(\d+)/);
  if (numA && numB && numA[1] === numB[1]) {
    // Same street number — check if street name overlaps
    const wordsA = new Set(na.split(' ').filter(w => w.length > 2 && isNaN(w)));
    const wordsB = nb.split(' ').filter(w => w.length > 2 && isNaN(w));
    if (wordsB.some(w => wordsA.has(w))) return true;
  }
  return false;
}

// Helper: find a matching account index with SAME address, or null
function pickSameSite(indices, oldAddr) {
  for (const idx of indices) {
    if (addrMatch(oldAddr, accounts[idx].address)) return idx;
  }
  return null; // no address match = different site, don't merge
}

let carriedForward = 0;
let mergedInto = 0;

existingData.forEach(d => {
  if (d.state !== 'KS' && d.state !== 'MO') return;
  if (!d.lat || !d.lng) return;
  if (!d.contacts || d.contacts.length === 0) return;

  const oldNorm = normName(d.company);
  const oldCore = coreName(d.company);
  const oldCity = (d.city || '').toLowerCase().trim();

  // All strategies require ADDRESS MATCH to merge — different address = different plant = separate pin

  // Strategy 1: exact normalized name + city + address match
  const key1 = oldNorm + '|' + oldCity;
  if (accountsByNormNameCity[key1]) {
    const samesite = pickSameSite(accountsByNormNameCity[key1], d.address);
    if (samesite !== null) {
      if (mergeContacts(accounts[samesite], d.contacts)) mergedInto++;
      return;
    }
  }

  // Strategy 2: core name + city + address match
  const key2 = oldCore + '|' + oldCity;
  if (accountsByCoreNameCity[key2] && accountsByCoreNameCity[key2].length <= 5) {
    const samesite = pickSameSite(accountsByCoreNameCity[key2], d.address);
    if (samesite !== null) {
      if (mergeContacts(accounts[samesite], d.contacts)) mergedInto++;
      return;
    }
  }

  // Strategy 3: normalized name only (any city, if unique) + address match
  if (accountsByNormName[oldNorm] && accountsByNormName[oldNorm].length === 1) {
    const samesite = pickSameSite(accountsByNormName[oldNorm], d.address);
    if (samesite !== null) {
      if (mergeContacts(accounts[samesite], d.contacts)) mergedInto++;
      return;
    }
  }

  // Strategy 4: word-overlap in same city + address match
  const cityAccounts = Object.entries(accountsByNormNameCity)
    .filter(([k]) => k.endsWith('|' + oldCity))
    .flatMap(([, indices]) => indices);

  const oldWords = new Set(oldNorm.split(' ').filter(w => w.length > 2));

  for (const idx of cityAccounts) {
    if (!addrMatch(d.address, accounts[idx].address)) continue;
    const exNorm = normName(accounts[idx].company);
    let score = 0;
    if (exNorm.includes(oldNorm) || oldNorm.includes(exNorm)) score += 5;
    const exWords = new Set(exNorm.split(' ').filter(w => w.length > 2));
    let shared = 0;
    oldWords.forEach(w => { if (exWords.has(w)) shared++; });
    score += shared * 2;
    const exCore = coreName(accounts[idx].company);
    if (oldCore === exCore) score += 3;
    if (score >= 4) {
      if (mergeContacts(accounts[idx], d.contacts)) mergedInto++;
      return;
    }
  }

  // No match at all — carry forward as standalone account
  const dedupKey = (d.company + '|' + (d.address || '')).toLowerCase();
  if (seenKeys.has(dedupKey)) return;
  seenKeys.add(dedupKey);

  const score = calcScoreOld(d);
  const grade = '?'; // assigned later via percentile quartiles
  // Old carried-forward accounts don't have confirmed 12-month revenue
  // Their sales field is stale historical data, not current — treat as prospect
  const status = 'prospect';

  const newAcct = {
    company: d.company,
    address: d.address || '',
    city: d.city || '',
    state: d.state,
    vertical: d.vertical || 'Other / Unknown',
    lat: d.lat,
    lng: d.lng,
    status,
    grade,
    score,
    rev012: 0,
    rev1224: 0,
    trend: 'unknown',
    upgradeReady: false,
    compCIJ: 0,
    compLaser: 0,
    prodLines: 0,
    oppNote: '',
    contacts: d.contacts || []
  };
  accounts.push(newAcct);

  // Register in lookup so subsequent old entries for same company+city can merge
  const newKey = normName(newAcct.company) + '|' + newAcct.city.toLowerCase().trim();
  if (!accountsByNormNameCity[newKey]) accountsByNormNameCity[newKey] = [];
  accountsByNormNameCity[newKey].push(accounts.length - 1);

  carriedForward++;
});

// ── 8c. Assign grades via percentile quartiles for even distribution ──
// Sort all scores descending, find quartile boundaries
const allScores = accounts.map(a => a.score).sort((a, b) => b - a);
const q25 = allScores[Math.floor(allScores.length * 0.25)]; // top 25% boundary
const q50 = allScores[Math.floor(allScores.length * 0.50)]; // top 50% boundary
const q75 = allScores[Math.floor(allScores.length * 0.75)]; // top 75% boundary

console.log(`Score quartile boundaries: A >= ${q25}, B >= ${q50}, C >= ${q75}, D < ${q75}`);
console.log(`Score range: ${allScores[0]} - ${allScores[allScores.length - 1]}`);

accounts.forEach(a => {
  if (a.score >= q25) a.grade = 'A';
  else if (a.score >= q50) a.grade = 'B';
  else if (a.score >= q75) a.grade = 'C';
  else a.grade = 'D';
});

// ── 8d. Health grade for active accounts ──
// Measures account HEALTH (how strong is this relationship?) vs the opportunity
// grade which measures POTENTIAL. Both grades show on active accounts.
//
// Health signals:
//   Revenue magnitude          — bigger spend = healthier
//   Revenue trend (YOY)        — growing vs declining
//   Revenue change %           — how much did they grow/shrink?
//   Tech breadth               — multiple tech types = deeper relationship
//   Contract customer           — service contract = sticky
//   Upgrade readiness           — engaged in refresh cycle
//   Break fix activity          — using our service = active relationship

const activeAccounts = accounts.filter(a => a.status === 'active');
activeAccounts.forEach(a => {
  let hp = 0;

  // Revenue magnitude (bigger wallet share = healthier)
  if (a.rev012 > 200000) hp += 25;
  else if (a.rev012 >= 100000) hp += 22;
  else if (a.rev012 >= 50000) hp += 18;
  else if (a.rev012 >= 20000) hp += 14;
  else if (a.rev012 >= 5000) hp += 8;
  else hp += 3;

  // Revenue trend
  if (a.trend === 'Rising') hp += 15;
  else if (a.trend === 'Declining') hp += 0;
  else hp += 7; // unknown/flat

  // Revenue change % (0-12 vs 12-24)
  if (a.rev1224 > 0) {
    const changePct = (a.rev012 - a.rev1224) / a.rev1224 * 100;
    if (changePct >= 20) hp += 15;       // strong growth
    else if (changePct >= 0) hp += 10;   // stable/slight growth
    else if (changePct >= -20) hp += 5;  // slight decline
    else hp += 0;                         // major decline
  } else {
    hp += 10; // new customer (no prior period), neutral
  }

  // Tech breadth (number of different VJ tech types)
  const techCount = a.vjEquip ? Object.keys(a.vjEquip).length : 0;
  if (techCount >= 4) hp += 15;
  else if (techCount >= 3) hp += 12;
  else if (techCount >= 2) hp += 8;
  else if (techCount >= 1) hp += 4;

  // Contract customer = sticky relationship
  if (a.isContractCustomer) hp += 10;

  // Upgrade readiness = engaged
  if (a.upgradeReady) hp += 8;

  // Break fix activity = active service relationship
  if (a.breakFixCount > 0) hp += 7;

  // Production lines (more lines = more opportunity for churn protection)
  if (a.prodLines >= 20) hp += 10;
  else if (a.prodLines >= 10) hp += 7;
  else if (a.prodLines >= 5) hp += 4;
  else if (a.prodLines >= 1) hp += 2;

  a.healthScore = hp;
});

// Assign health grades via percentile quartiles (among active accounts only)
if (activeAccounts.length > 0) {
  const healthScores = activeAccounts.map(a => a.healthScore).sort((a, b) => b - a);
  const hq25 = healthScores[Math.floor(healthScores.length * 0.25)];
  const hq50 = healthScores[Math.floor(healthScores.length * 0.50)];
  const hq75 = healthScores[Math.floor(healthScores.length * 0.75)];

  console.log(`Health quartile boundaries: A >= ${hq25}, B >= ${hq50}, C >= ${hq75}, D < ${hq75}`);

  activeAccounts.forEach(a => {
    if (a.healthScore >= hq25) a.healthGrade = 'A';
    else if (a.healthScore >= hq50) a.healthGrade = 'B';
    else if (a.healthScore >= hq75) a.healthGrade = 'C';
    else a.healthGrade = 'D';
  });

  console.log(`Health grades: A:${activeAccounts.filter(a=>a.healthGrade==='A').length} B:${activeAccounts.filter(a=>a.healthGrade==='B').length} C:${activeAccounts.filter(a=>a.healthGrade==='C').length} D:${activeAccounts.filter(a=>a.healthGrade==='D').length}`);
}

console.log(`Processed ${accounts.length} accounts (${accounts.length - carriedForward} from Excel/merged, ${carriedForward} carried forward, ${mergedInto} old entries merged into Excel accounts)`);
console.log(`  Active: ${accounts.filter(a=>a.status==='active').length}`);
console.log(`  Lapsed: ${accounts.filter(a=>a.status==='lapsed').length}`);
console.log(`  Dormant: ${accounts.filter(a=>a.status==='dormant').length}`);
console.log(`  Prospect: ${accounts.filter(a=>a.status==='prospect').length}`);
console.log(`  A-grade: ${accounts.filter(a=>a.grade==='A').length}`);
console.log(`  B-grade: ${accounts.filter(a=>a.grade==='B').length}`);
console.log(`  C-grade: ${accounts.filter(a=>a.grade==='C').length}`);
console.log(`  D-grade: ${accounts.filter(a=>a.grade==='D').length}`);

// Sort: active first, then lapsed, then by grade
accounts.sort((a, b) => {
  const statusOrder = { active: 0, lapsed: 1, dormant: 2, prospect: 3 };
  const gradeOrder = { A: 0, B: 1, C: 2, D: 3 };
  if (statusOrder[a.status] !== statusOrder[b.status]) return statusOrder[a.status] - statusOrder[b.status];
  if (gradeOrder[a.grade] !== gradeOrder[b.grade]) return gradeOrder[a.grade] - gradeOrder[b.grade];
  return a.company.localeCompare(b.company);
});

const RAW_DATA_JSON = JSON.stringify(accounts);

// ── 9. Generate HTML files ──

function getVerticals(data) {
  const verts = {};
  data.forEach(d => { verts[d.vertical] = (verts[d.vertical]||0)+1; });
  return Object.entries(verts).sort((a,b) => b[1]-a[1]).slice(0, 15).map(v => v[0]);
}

const topVerticals = getVerticals(accounts);

// ── DESKTOP HTML ──
const desktopHTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Videojet Territory Intelligence Map</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"><\/script>
<style>
:root {
  --bg: #0a0e14;
  --surface: #111820;
  --surface2: #161e28;
  --border: #1e2d3d;
  --border2: #243344;
  --accent: #00d4ff;
  --accent2: #0099cc;
  --text: #c8d8e8;
  --text-dim: #5a7a94;
  --text-bright: #e8f4ff;
  --active-color: #FFD700;
  --active-glow: rgba(255,215,0,0.4);
  --lapsed-color: #B71C1C;
  --lapsed-glow: rgba(183,28,28,0.5);
  --dormant-color: #F48FB1;
  --dormant-glow: rgba(244,143,177,0.4);
  --a-color: #2ECC71;
  --a-glow: rgba(46,204,113,0.35);
  --b-color: #3498DB;
  --b-glow: rgba(52,152,219,0.35);
  --c-color: #FF8C00;
  --c-glow: rgba(255,140,0,0.35);
  --d-color: #95A5A6;
  --d-glow: rgba(149,165,166,0.35);
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'IBM Plex Sans',sans-serif; background:var(--bg); color:var(--text); height:100vh; display:flex; flex-direction:column; overflow:hidden; }

/* HEADER */
header { background:var(--surface); border-bottom:1px solid var(--border); padding:0 20px; height:56px; display:flex; align-items:center; justify-content:space-between; flex-shrink:0; position:relative; z-index:1000; }
header::after { content:''; position:absolute; bottom:0; left:0; right:0; height:1px; background:linear-gradient(90deg,transparent,var(--accent),transparent); opacity:0.4; }
.logo { font-family:'Rajdhani',sans-serif; font-size:20px; font-weight:700; letter-spacing:3px; color:var(--text-bright); text-transform:uppercase; display:flex; align-items:center; gap:10px; }
.logo-dot { width:8px; height:8px; background:var(--accent); border-radius:50%; box-shadow:0 0 12px var(--accent); animation:pulse 2s ease-in-out infinite; }
@keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:0.5;transform:scale(0.7)} }
.header-stats { display:flex; gap:24px; font-family:'IBM Plex Mono',monospace; font-size:11px; }
.stat { display:flex; flex-direction:column; align-items:center; gap:2px; }
.stat-value { font-size:16px; font-weight:500; color:var(--accent); }
.stat-label { color:var(--text-dim); text-transform:uppercase; letter-spacing:1px; }
.header-actions { display:flex; gap:8px; }
.hdr-btn { padding:6px 12px; border-radius:4px; font-family:'IBM Plex Mono',monospace; font-size:10px; cursor:pointer; border:1px solid var(--border2); background:var(--surface2); color:var(--text-dim); transition:all 0.15s; text-transform:uppercase; letter-spacing:0.5px; }
.hdr-btn:hover { border-color:var(--accent); color:var(--accent); }

/* LAYOUT */
.main { display:flex; flex:1; overflow:hidden; }

/* SIDEBAR */
.sidebar { width:300px; background:var(--surface); border-right:1px solid var(--border); display:flex; flex-direction:column; flex-shrink:0; overflow:hidden; }
.sidebar-section { padding:14px 16px; border-bottom:1px solid var(--border); }
.section-label { font-family:'IBM Plex Mono',monospace; font-size:9px; text-transform:uppercase; letter-spacing:2px; color:var(--text-dim); margin-bottom:10px; }
.search-wrap { position:relative; }
.search-wrap input { width:100%; background:var(--bg); border:1px solid var(--border2); border-radius:4px; color:var(--text-bright); padding:8px 12px 8px 32px; font-family:'IBM Plex Sans',sans-serif; font-size:13px; outline:none; transition:border-color 0.2s; }
.search-wrap input:focus { border-color:var(--accent); }
.search-wrap input::placeholder { color:var(--text-dim); }
.search-icon { position:absolute; left:10px; top:50%; transform:translateY(-50%); color:var(--text-dim); font-size:14px; pointer-events:none; }
.filter-group { margin-bottom:12px; }
.filter-label { font-size:10px; color:var(--text-dim); text-transform:uppercase; letter-spacing:1px; margin-bottom:6px; display:block; }
.filter-pills { display:flex; flex-wrap:wrap; gap:5px; }
.pill { padding:4px 10px; border-radius:3px; font-size:11px; font-family:'IBM Plex Mono',monospace; cursor:pointer; border:1px solid var(--border2); background:var(--bg); color:var(--text-dim); transition:all 0.15s; user-select:none; letter-spacing:0.5px; }
.pill:hover { border-color:var(--accent); color:var(--text); }
.pill.active { background:var(--accent); border-color:var(--accent); color:#000; font-weight:600; }
.pill.grade-a.active { background:var(--a-color); border-color:var(--a-color); }
.pill.grade-b.active { background:var(--b-color); border-color:var(--b-color); }
.pill.grade-c.active { background:var(--c-color); border-color:var(--c-color); }
.pill.status-active.active { background:var(--active-color); border-color:var(--active-color); }
.pill.status-lapsed.active { background:var(--lapsed-color); border-color:var(--lapsed-color); }
.pill.status-dormant.active { background:var(--dormant-color); border-color:var(--dormant-color); }
.results-header { display:flex; align-items:center; justify-content:space-between; padding:10px 16px 8px; border-bottom:1px solid var(--border); }
.results-count { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--text-dim); }
.results-count span { color:var(--accent); font-weight:500; }
.clear-btn { font-size:10px; color:var(--text-dim); cursor:pointer; text-transform:uppercase; letter-spacing:1px; padding:2px 6px; border:1px solid var(--border2); border-radius:3px; transition:all 0.15s; }
.clear-btn:hover { color:var(--accent); border-color:var(--accent); }
.results-list { flex:1; overflow-y:auto; padding:4px 0; }
.results-list::-webkit-scrollbar { width:4px; }
.results-list::-webkit-scrollbar-track { background:transparent; }
.results-list::-webkit-scrollbar-thumb { background:var(--border2); border-radius:2px; }
.result-item { padding:9px 16px; cursor:pointer; border-bottom:1px solid var(--border); transition:background 0.1s; position:relative; }
.result-item:hover { background:var(--surface2); }
.result-item.selected { background:rgba(0,212,255,0.06); }
.result-item::before { content:''; position:absolute; left:0; top:0; bottom:0; width:3px; }
.result-item.status-active::before { background:var(--active-color); }
.result-item.status-lapsed::before { background:var(--lapsed-color); }
.result-item.status-dormant::before { background:var(--dormant-color); }
.result-item.grade-A.status-prospect::before { background:var(--a-color); }
.result-item.grade-B.status-prospect::before { background:var(--b-color); }
.result-item.grade-C.status-prospect::before { background:var(--c-color); }
.result-item.grade-D.status-prospect::before { background:var(--d-color); }
.result-company { font-size:12px; font-weight:500; color:var(--text-bright); line-height:1.3; margin-bottom:3px; }
.result-meta { display:flex; align-items:center; gap:8px; font-family:'IBM Plex Mono',monospace; font-size:10px; color:var(--text-dim); }
.result-grade { font-weight:600; padding:1px 5px; border-radius:2px; font-size:10px; }
.result-grade.A { background:rgba(46,204,113,0.15); color:var(--a-color); }
.result-grade.B { background:rgba(52,152,219,0.15); color:var(--b-color); }
.result-grade.C { background:rgba(255,140,0,0.15); color:var(--c-color); }
.result-grade.D { background:rgba(149,165,166,0.15); color:var(--d-color); }
.result-status-badge { padding:1px 5px; border-radius:2px; font-size:9px; font-weight:600; letter-spacing:0.5px; }
.result-status-badge.active { background:rgba(255,215,0,0.15); color:var(--active-color); }
.result-status-badge.lapsed { background:rgba(183,28,28,0.15); color:var(--lapsed-color); }
.result-status-badge.dormant { background:rgba(244,143,177,0.15); color:var(--dormant-color); }
.result-star { color:#FFD700; margin-right:4px; }
.result-ring { color:#FF4444; margin-right:4px; font-size:8px; }

/* MAP */
#map { flex:1; position:relative; }
.leaflet-tile { filter:none; }
.leaflet-control-zoom a { background:var(--surface)!important; color:var(--text)!important; border-color:var(--border)!important; }
.leaflet-control-zoom a:hover { background:var(--surface2)!important; color:var(--accent)!important; }
.leaflet-control-attribution { display:none; }

/* Markers */
.map-marker { width:14px; height:14px; border-radius:50%; border:2px solid rgba(255,255,255,0.3); cursor:pointer; transition:transform 0.15s; position:relative; }
.map-marker:hover { transform:scale(1.5); z-index:999!important; }
.map-marker.status-active { background:var(--active-color); box-shadow:0 0 10px var(--active-glow); border:2px solid rgba(255,215,0,0.6); }
.map-marker.status-lapsed { background:var(--lapsed-color); box-shadow:0 0 12px var(--lapsed-glow); border:2px solid rgba(183,28,28,0.7); }
.map-marker.status-dormant { background:var(--dormant-color); box-shadow:0 0 12px var(--dormant-glow); border:2px solid rgba(244,143,177,0.6); }
.map-marker.grade-A { background:var(--a-color); box-shadow:0 0 8px var(--a-glow); }
.map-marker.grade-B { background:var(--b-color); box-shadow:0 0 8px var(--b-glow); }
.map-marker.grade-C { background:var(--c-color); box-shadow:0 0 8px var(--c-glow); }
.map-marker.grade-D { background:var(--d-color); box-shadow:0 0 6px var(--d-glow); }
.map-marker.selected { transform:scale(1.6); z-index:999!important; border:2px solid white; }
.map-marker .star-overlay { position:absolute; top:-8px; right:-8px; font-size:10px; pointer-events:none; }
.map-marker .ring-overlay { position:absolute; top:-2px; left:-2px; right:-2px; bottom:-2px; border-radius:50%; border:2px solid #FF4444; pointer-events:none; animation:ringPulse 1.5s infinite; }
@keyframes ringPulse { 0%,100%{opacity:1} 50%{opacity:0.4} }

/* Cluster */
.my-cluster { border-radius:50%; background:rgba(17,24,32,0.92); border:2px solid #243344; display:flex; align-items:center; justify-content:center; font-family:'IBM Plex Mono',monospace; font-weight:600; color:#e8f4ff; box-shadow:0 2px 10px rgba(0,0,0,0.5); cursor:pointer; }

/* Popup */
.leaflet-popup-content-wrapper { background:var(--surface)!important; border:1px solid var(--border2)!important; border-radius:6px!important; box-shadow:0 8px 32px rgba(0,0,0,0.6)!important; padding:0!important; }
.leaflet-popup-tip { background:var(--surface)!important; }
.leaflet-popup-close-button { color:var(--text-dim)!important; font-size:18px!important; padding:6px 10px!important; }
.leaflet-popup-close-button:hover { color:var(--text-bright)!important; }
.leaflet-popup-content { margin:0!important; padding:0!important; width:auto!important; }
.popup-inner { padding:14px 16px; min-width:260px; max-width:340px; }
.popup-company { font-family:'Rajdhani',sans-serif; font-size:15px; font-weight:600; color:var(--text-bright); margin-bottom:8px; line-height:1.2; letter-spacing:0.5px; }
.popup-row { display:flex; align-items:center; gap:8px; font-size:11px; color:var(--text-dim); margin-bottom:4px; font-family:'IBM Plex Mono',monospace; }
.popup-row strong { color:var(--text); font-weight:500; }
.popup-badges { display:flex; gap:6px; margin-top:10px; flex-wrap:wrap; }
.badge { padding:3px 8px; border-radius:3px; font-size:10px; font-family:'IBM Plex Mono',monospace; font-weight:600; letter-spacing:0.5px; }
.badge-a { background:rgba(46,204,113,0.15); color:var(--a-color); border:1px solid rgba(46,204,113,0.3); }
.badge-b { background:rgba(52,152,219,0.15); color:var(--b-color); border:1px solid rgba(52,152,219,0.3); }
.badge-c { background:rgba(255,140,0,0.15); color:var(--c-color); border:1px solid rgba(255,140,0,0.3); }
.badge-d { background:rgba(149,165,166,0.15); color:var(--d-color); border:1px solid rgba(149,165,166,0.3); }
.badge-active { background:rgba(255,215,0,0.15); color:var(--active-color); border:1px solid rgba(255,215,0,0.3); }
.badge-lapsed { background:rgba(183,28,28,0.15); color:var(--lapsed-color); border:1px solid rgba(183,28,28,0.3); }
.badge-dormant { background:rgba(244,143,177,0.15); color:var(--dormant-color); border:1px solid rgba(244,143,177,0.3); }
.badge-vert { background:rgba(0,212,255,0.08); color:var(--accent); border:1px solid rgba(0,212,255,0.2); }
.popup-details { margin-top:8px; padding-top:8px; border-top:1px solid var(--border); font-family:'IBM Plex Mono',monospace; font-size:10px; }
.popup-detail-row { display:flex; justify-content:space-between; margin-bottom:3px; color:var(--text-dim); }
.popup-detail-row .val { color:var(--text); font-weight:500; }
.popup-detail-row .val.green { color:var(--a-color); }
.popup-detail-row .val.red { color:#FF6B6B; }
.popup-note { margin-top:6px; padding:6px 8px; background:rgba(0,212,255,0.05); border:1px solid var(--border); border-radius:3px; font-size:10px; color:var(--text); font-family:'IBM Plex Mono',monospace; }
.popup-actions { margin-top:10px; display:flex; gap:6px; flex-wrap:wrap; }
.popup-btn { display:inline-flex; align-items:center; gap:5px; padding:6px 10px; border-radius:4px; cursor:pointer; font-family:'IBM Plex Mono',monospace; font-size:10px; font-weight:600; border:1px solid var(--border2); color:var(--text-dim); background:var(--bg); transition:all 0.15s; }
.popup-btn:hover { border-color:var(--accent); color:var(--accent); }
.popup-btn.primary { border-color:var(--accent); color:var(--accent); background:rgba(0,212,255,0.08); }

/* Contacts panel */
.contacts-panel { position:absolute; top:0; right:0; bottom:0; width:380px; background:var(--surface); border-left:1px solid var(--border2); z-index:2000; display:flex; flex-direction:column; transform:translateX(100%); transition:transform 0.25s ease; box-shadow:-8px 0 32px rgba(0,0,0,0.4); }
.contacts-panel.open { transform:translateX(0); }
.cp-header { padding:16px 18px; border-bottom:1px solid var(--border); display:flex; align-items:flex-start; justify-content:space-between; flex-shrink:0; }
.cp-company { font-family:'Rajdhani',sans-serif; font-size:15px; font-weight:600; color:var(--text-bright); letter-spacing:0.5px; line-height:1.2; flex:1; padding-right:12px; }
.cp-close { color:var(--text-dim); cursor:pointer; font-size:20px; line-height:1; flex-shrink:0; transition:color 0.15s; }
.cp-close:hover { color:var(--text-bright); }
.cp-list { flex:1; overflow-y:auto; padding:8px 0; }
.cp-list::-webkit-scrollbar { width:4px; }
.cp-list::-webkit-scrollbar-thumb { background:var(--border2); border-radius:2px; }
.cp-contact { padding:10px 18px; border-bottom:1px solid var(--border); position:relative; }
.cp-contact::before { content:''; position:absolute; left:0; top:0; bottom:0; width:3px; }
.cp-contact.equip::before { background:var(--accent); }
.cp-contact.fit-A::before { background:var(--a-color); }
.cp-contact.fit-B::before { background:var(--b-color); }
.cp-contact.fit-other::before { background:var(--text-dim); opacity:0.4; }
.cp-name { font-size:13px; font-weight:600; color:var(--text-bright); margin-bottom:2px; display:flex; align-items:center; gap:7px; }
.cp-dm-badge { font-family:'IBM Plex Mono',monospace; font-size:8px; font-weight:700; padding:1px 5px; border-radius:2px; letter-spacing:0.5px; background:rgba(255,215,0,0.15); color:var(--active-color); border:1px solid rgba(255,215,0,0.3); }
.cp-equip-badge { font-family:'IBM Plex Mono',monospace; font-size:8px; font-weight:700; padding:1px 5px; border-radius:2px; letter-spacing:0.5px; background:rgba(0,212,255,0.15); color:var(--accent); border:1px solid rgba(0,212,255,0.3); }
.cp-title { font-size:11px; color:var(--text); margin-bottom:4px; }
.cp-phone-row { font-size:11px; margin-top:2px; display:flex; align-items:center; gap:5px; }
.cp-phone { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--accent); text-decoration:none; }
.cp-phone:hover { text-decoration:underline; }
.cp-email { font-family:'IBM Plex Mono',monospace; font-size:10px; color:var(--accent); text-decoration:none; }
.cp-email:hover { text-decoration:underline; }
.cp-nophone { font-size:10px; color:var(--text-dim); font-style:italic; margin-top:3px; font-family:'IBM Plex Mono',monospace; }
.cp-empty { padding:24px 18px; color:var(--text-dim); font-family:'IBM Plex Mono',monospace; font-size:11px; text-align:center; }

/* Edit modal */
.edit-modal-overlay { position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.6); z-index:3000; display:none; align-items:center; justify-content:center; }
.edit-modal-overlay.open { display:flex; }
.edit-modal { background:var(--surface); border:1px solid var(--border2); border-radius:8px; padding:20px; min-width:320px; max-width:400px; box-shadow:0 16px 48px rgba(0,0,0,0.5); }
.edit-modal h3 { font-family:'Rajdhani',sans-serif; font-size:16px; color:var(--text-bright); margin-bottom:14px; }
.edit-row { display:flex; align-items:center; gap:10px; margin-bottom:10px; }
.edit-row label { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--text-dim); min-width:80px; }
.edit-row select, .edit-row input, .edit-row textarea { background:var(--bg); border:1px solid var(--border2); border-radius:4px; color:var(--text-bright); padding:6px 10px; font-family:'IBM Plex Sans',sans-serif; font-size:12px; outline:none; flex:1; }
.edit-row textarea { resize:vertical; min-height:60px; }
.edit-row select:focus, .edit-row input:focus, .edit-row textarea:focus { border-color:var(--accent); }
.edit-btns { display:flex; gap:8px; margin-top:14px; justify-content:flex-end; }
.edit-btn { padding:7px 14px; border-radius:4px; font-family:'IBM Plex Mono',monospace; font-size:11px; cursor:pointer; border:1px solid var(--border2); background:var(--bg); color:var(--text-dim); transition:all 0.15s; }
.edit-btn:hover { border-color:var(--accent); color:var(--accent); }
.edit-btn.save { border-color:var(--a-color); color:var(--a-color); background:rgba(46,204,113,0.08); }
.edit-btn.danger { border-color:#FF6B6B; color:#FF6B6B; }
.edit-btn.danger:hover { background:rgba(255,107,107,0.1); }

/* Add Account modal */
.add-modal-overlay { position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.6); z-index:3000; display:none; align-items:center; justify-content:center; }
.add-modal-overlay.open { display:flex; }

/* Legend */
.map-legend { position:absolute; bottom:24px; right:16px; background:var(--surface); border:1px solid var(--border2); border-radius:6px; padding:12px 14px; z-index:1000; min-width:180px; }
.legend-title { font-family:'IBM Plex Mono',monospace; font-size:9px; text-transform:uppercase; letter-spacing:2px; color:var(--text-dim); margin-bottom:10px; }
.legend-item { display:flex; align-items:center; gap:8px; margin-bottom:7px; font-size:11px; color:var(--text); }
.legend-dot { width:11px; height:11px; border-radius:50%; flex-shrink:0; }
</style>
</head>
<body>

<header>
  <div class="logo"><div class="logo-dot"></div>VIDEOJET TERRITORY INTEL</div>
  <div class="header-stats">
    <div class="stat"><div class="stat-value" id="hdr-total" style="color:var(--text-bright)">0</div><div class="stat-label">Accounts</div></div>
    <div class="stat"><div class="stat-value" id="hdr-active" style="color:var(--active-color)">0</div><div class="stat-label">Active</div></div>
    <div class="stat"><div class="stat-value" id="hdr-lapsed" style="color:var(--lapsed-color)">0</div><div class="stat-label">Lapsed</div></div>
    <div class="stat"><div class="stat-value" id="hdr-dormant" style="color:var(--dormant-color)">0</div><div class="stat-label">Dormant</div></div>
    <div class="stat"><div class="stat-value" id="hdr-prospect" style="color:var(--accent)">0</div><div class="stat-label">Prospects</div></div>
  </div>
  <div class="header-actions">
    <div class="hdr-btn" onclick="openAddAccount()">+ Add Account</div>
    <div class="hdr-btn" onclick="exportEdits()">Export Edits</div>
  </div>
</header>

<div class="main">
  <div class="sidebar">
    <div class="sidebar-section">
      <div class="section-label">Search</div>
      <div class="search-wrap">
        <span class="search-icon">&#x2315;</span>
        <input type="text" id="search" placeholder="Company, city, vertical...">
      </div>
    </div>
    <div class="sidebar-section">
      <div class="section-label">Filters</div>
      <div class="filter-group">
        <span class="filter-label">State</span>
        <div class="filter-pills" id="state-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="MO">MO</div>
          <div class="pill" data-val="KS">KS</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Grade</span>
        <div class="filter-pills" id="grade-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill grade-a" data-val="A">A</div>
          <div class="pill grade-b" data-val="B">B</div>
          <div class="pill grade-c" data-val="C">C</div>
          <div class="pill" data-val="D">D</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Status</span>
        <div class="filter-pills" id="status-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill status-active" data-val="active">Active</div>
          <div class="pill status-lapsed" data-val="lapsed">Lapsed</div>
          <div class="pill status-dormant" data-val="dormant">Dormant</div>
          <div class="pill" data-val="prospect">Prospect</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Vertical / Industry</span>
        <div class="filter-pills" id="vert-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="Food">Food</div>
          <div class="pill" data-val="Beverage">Beverage</div>
          <div class="pill" data-val="PetFood">Pet Food</div>
          <div class="pill" data-val="Pharma">Pharma</div>
          <div class="pill" data-val="Packaging">Packaging</div>
          <div class="pill" data-val="Aero">Aero/Defense</div>
          <div class="pill" data-val="Auto">Auto/Trans</div>
          <div class="pill" data-val="Chemical">Chemical</div>
          <div class="pill" data-val="Plastics">Plastics</div>
          <div class="pill" data-val="Building">Building</div>
          <div class="pill" data-val="Industrial">Industrial</div>
          <div class="pill" data-val="Electronics">Electronics</div>
          <div class="pill" data-val="Consumer">CPG</div>
          <div class="pill" data-val="Graphics">Graphics</div>
          <div class="pill" data-val="Other">Other</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Health Grade</span>
        <div class="filter-pills" id="health-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill grade-a" data-val="A">A</div>
          <div class="pill grade-b" data-val="B">B</div>
          <div class="pill grade-c" data-val="C">C</div>
          <div class="pill" data-val="D">D</div>
          <div class="pill" data-val="none">N/A</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Upgrade Ready</span>
        <div class="filter-pills" id="upgrade-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="Y">Yes</div>
          <div class="pill" data-val="N">No</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Contacts</span>
        <div class="filter-pills" id="contacts-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="0">0</div>
          <div class="pill" data-val="1">1</div>
          <div class="pill" data-val="2+">2+</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Open Opportunities</span>
        <div class="filter-pills" id="opps-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="Y">Has Opps</div>
          <div class="pill" data-val="N">None</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Revenue (12mo)</span>
        <div class="filter-pills" id="rev-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="0">$0</div>
          <div class="pill" data-val="1-10k">&lt;$10K</div>
          <div class="pill" data-val="10-50k">$10-50K</div>
          <div class="pill" data-val="50-100k">$50-100K</div>
          <div class="pill" data-val="100k+">$100K+</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Starred</span>
        <div class="filter-pills" id="starred-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="Y">&#x2B50; Starred</div>
          <div class="pill" data-val="N">Not Starred</div>
        </div>
      </div>
      <div class="filter-group">
        <span class="filter-label">Comments</span>
        <div class="filter-pills" id="commented-filter">
          <div class="pill active" data-val="ALL">All</div>
          <div class="pill" data-val="Y">Has Comments</div>
          <div class="pill" data-val="N">No Comments</div>
        </div>
      </div>
    </div>
    <div class="results-header">
      <div class="results-count">Showing <span id="result-count">0</span> accounts</div>
      <div style="display:flex;gap:6px">
        <div class="clear-btn" onclick="openRulesPanel()">RULES</div>
        <div class="clear-btn" id="clear-filters">RESET</div>
      </div>
    </div>
    <div class="results-list" id="results-list"></div>
  </div>

  <div id="map">
    <div class="contacts-panel" id="contacts-panel">
      <div class="cp-header">
        <div class="cp-company" id="cp-company-name"></div>
        <div class="cp-close" onclick="closeContacts()">&#x2715;</div>
      </div>
      <div class="cp-list" id="cp-contact-list"></div>
    </div>
    <div class="map-legend">
      <div class="legend-title">Legend</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--active-color);box-shadow:0 0 6px var(--active-glow)"></div> Active Customer</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--lapsed-color);box-shadow:0 0 6px var(--lapsed-glow)"></div> Lapsed Customer</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--dormant-color);box-shadow:0 0 6px var(--dormant-glow)"></div> Dormant (has equip)</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--a-color);box-shadow:0 0 6px var(--a-glow)"></div> A-Grade Prospect</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--b-color);box-shadow:0 0 6px var(--b-glow)"></div> B-Grade Prospect</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--c-color);box-shadow:0 0 6px var(--c-glow)"></div> C-Grade Prospect</div>
      <div class="legend-item"><div class="legend-dot" style="background:var(--d-color);box-shadow:0 0 6px var(--d-glow)"></div> D-Grade Prospect</div>
      <div class="legend-item" style="margin-top:6px;border-top:1px solid var(--border);padding-top:6px"><span style="font-size:12px">&#x2B50;</span> Starred Account</div>
      <div class="legend-item"><div style="width:11px;height:11px;border-radius:50%;border:2px solid #FF4444"></div> Active Trial/Opp</div>
    </div>
  </div>
</div>

<!-- Edit Modal -->
<div class="edit-modal-overlay" id="edit-modal">
  <div class="edit-modal">
    <h3 id="edit-modal-title">Edit Account</h3>
    <div class="edit-row"><label>Star</label><select id="edit-star"><option value="0">No</option><option value="1">&#x2B50; Yes</option></select></div>
    <div class="edit-row"><label>Trial/Opp</label><select id="edit-ring"><option value="0">No</option><option value="1">&#x1F534; Active Trial</option></select></div>
    <div class="edit-row"><label>Grade</label><select id="edit-grade"><option value="">Auto</option><option value="A">A</option><option value="B">B</option><option value="C">C</option><option value="D">D</option></select></div>
    <div class="edit-row"><label>Status</label><select id="edit-status"><option value="">Auto</option><option value="active">Active Customer</option><option value="lapsed">Lapsed Customer</option><option value="dormant">Dormant</option><option value="prospect">Prospect</option></select></div>
    <div class="edit-row"><label>Note</label><textarea id="edit-note" placeholder="Add a note..."></textarea></div>
    <div class="edit-btns">
      <div class="edit-btn danger" id="edit-delete-btn">Delete</div>
      <div class="edit-btn" onclick="closeEditModal()">Cancel</div>
      <div class="edit-btn save" onclick="saveEdit()">Save</div>
    </div>
  </div>
</div>

<!-- Add Account Modal -->
<div class="add-modal-overlay" id="add-modal">
  <div class="edit-modal">
    <h3>Add Account</h3>
    <div class="edit-row"><label>Name</label><input id="add-name" type="text" placeholder="Company name"></div>
    <div class="edit-row"><label>Address</label><input id="add-address" type="text" placeholder="Street address"></div>
    <div class="edit-row"><label>City</label><input id="add-city" type="text" placeholder="City"></div>
    <div class="edit-row"><label>State</label><select id="add-state"><option value="KS">KS</option><option value="MO">MO</option></select></div>
    <div class="edit-row"><label>Grade</label><select id="add-grade"><option value="C">C</option><option value="A">A</option><option value="B">B</option><option value="D">D</option></select></div>
    <div class="edit-row"><label>Note</label><textarea id="add-note" placeholder="Note..."></textarea></div>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text-dim);margin-bottom:10px;">Click on the map to set the pin location, or enter coords:</div>
    <div class="edit-row"><label>Lat</label><input id="add-lat" type="number" step="0.0001" placeholder="38.0000"></div>
    <div class="edit-row"><label>Lng</label><input id="add-lng" type="number" step="0.0001" placeholder="-96.0000"></div>
    <div class="edit-btns">
      <div class="edit-btn" onclick="closeAddModal()">Cancel</div>
      <div class="edit-btn save" onclick="saveNewAccount()">Add</div>
    </div>
  </div>
</div>

<!-- Rules Panel -->
<div class="edit-modal-overlay" id="rules-panel">
  <div class="edit-modal" style="max-width:520px;max-height:85vh;overflow-y:auto">
    <h3>Rules &amp; Scoring</h3>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--text);line-height:1.7">

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:2px;margin:12px 0 6px">Account Statuses</div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--active-color)"></span> <strong>Active (Gold)</strong> &mdash; Has VJ revenue in the last 12 months.</div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--lapsed-color)"></span> <strong>Lapsed (Red)</strong> &mdash; Salesforce Type = &ldquo;Active Customer&rdquo; but zero revenue last 12 months. CRM still considers them active, but they stopped buying.</div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--dormant-color)"></span> <strong>Dormant (Pink)</strong> &mdash; Has VJ equipment in ERP but no orders last 12 months, and CRM types them as Prospect. They bought systems from us but fell off. Also includes accounts with revenue only 12-24 months ago.</div>
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--a-color)"></span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--b-color)"></span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--c-color)"></span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--d-color)"></span> <strong>Prospect</strong> &mdash; No recent VJ revenue or equipment. Color = opportunity grade.</div>

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:2px;margin:16px 0 6px">Opportunity Grade (all accounts)</div>
      <div style="margin-bottom:6px"><em>&ldquo;How likely does this site need marking/coding?&rdquo;</em></div>
      <div style="margin-bottom:4px;color:var(--text-bright)">Core signals (~70%):</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Vertical fit</strong> (max 30) &mdash; Food/pharma/bev 28-30, pkg/chem 18-22, auto/aero 15, graphics 8</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Plant size</strong> (max 25) &mdash; 500+ emp=25, 200+=22, 100+=18, 50+=14, 20+=9</div>
      <div style="padding-left:12px;margin-bottom:6px"><strong>Company revenue</strong> (max 15) &mdash; $1B+=15, $100M+=11, $10M+=8, $1M+=3</div>
      <div style="margin-bottom:4px;color:var(--text-bright)">Bonus signals (no penalty for missing data):</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Competitive installs</strong> (max 10) &mdash; 10+=10, 5+=7, 1+=4</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>VJ installed base</strong> (max 8) &mdash; 10+=8, 5+=5, 1+=3</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Current VJ revenue</strong> (max 6) &mdash; $100K+=6, $50K+=5, $10K+=3</div>
      <div style="padding-left:12px;margin-bottom:6px"><strong>Upgrade readiness</strong> &mdash; 4pts if Y</div>
      <div style="color:var(--text-dim)">Top 25% = A, next 25% = B, next 25% = C, bottom 25% = D</div>

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:2px;margin:16px 0 6px">Health Grade (active accounts only)</div>
      <div style="margin-bottom:6px"><em>&ldquo;How strong is this customer relationship?&rdquo;</em></div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Revenue size</strong> (max 25) &mdash; $200K+=25, $100K+=22, $50K+=18, $20K+=14, $5K+=8</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Revenue trend</strong> (max 15) &mdash; Rising=15, Unknown=7, Declining=0</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>YOY change %</strong> (max 15) &mdash; +20%=15, stable=10, slight decline=5, major decline=0</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Tech breadth</strong> (max 15) &mdash; 4+ types=15, 3=12, 2=8, 1=4</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Service contract</strong> &mdash; 10pts if yes</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Upgrade readiness</strong> &mdash; 8pts if Y</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Break-fix activity</strong> &mdash; 7pts if any</div>
      <div style="padding-left:12px;margin-bottom:6px"><strong>Production lines</strong> (max 10) &mdash; 20+=10, 10+=7, 5+=4, 1+=2</div>
      <div style="color:var(--text-dim)">Graded among active accounts only. A = thriving, D = at-risk.</div>

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:2px;margin:16px 0 6px">Reading the two grades together</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Opp A + Health A</strong> &mdash; Strong account, protect it</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Opp A + Health D</strong> &mdash; Big site declining, needs attention</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>Opp D + Health A</strong> &mdash; Small but loyal, low priority</div>
      <div style="padding-left:12px;margin-bottom:2px"><strong>High Opp prospect</strong> &mdash; High-value site not buying from us yet</div>
    </div>
    <div class="edit-btns" style="margin-top:16px">
      <div class="edit-btn" onclick="closeRulesPanel()">Close</div>
    </div>
  </div>
</div>

<script>
const RAW_DATA = ${RAW_DATA_JSON};

document.addEventListener("DOMContentLoaded", function() {

// ── LocalStorage edits ──
const LS_KEY = 'vj_territory_edits';
function loadEdits() { try { return JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch(e) { return {}; } }
function saveEdits(edits) { localStorage.setItem(LS_KEY, JSON.stringify(edits)); }
function getEditKey(d) { return (d.company + '|' + (d.address || '')).toLowerCase(); }

// Apply edits to data
function getEffective(d, idx) {
  const edits = loadEdits();
  const key = getEditKey(d);
  const e = edits[key] || {};
  return {
    ...d,
    grade: e.grade || d.grade,
    status: e.status || d.status,
    star: !!e.star,
    ring: !!e.ring,
    note: e.note || '',
    deleted: !!e.deleted
  };
}

// ── State ──
let filters = { state:'ALL', grade:'ALL', status:'ALL', vert:'ALL', health:'ALL', upgrade:'ALL', contacts:'ALL', opps:'ALL', rev:'ALL', starred:'ALL', commented:'ALL', search:'' };
let markers = [];
let selectedIdx = null;

// ── Map init ──
const map = L.map('map', { center:[38.8,-96.5], zoom:6, zoomControl:true });
L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', { maxZoom:19, subdomains:'abcd' }).addTo(map);

const plainGroup = L.layerGroup().addTo(map);



function formatSales(n) {
  if (!n) return '$0';
  if (n >= 1000000) return '$' + (n/1000000).toFixed(1) + 'M';
  if (n >= 1000) return '$' + Math.round(n/1000) + 'K';
  return '$' + n;
}

function matchesFilters(d) {
  const eff = getEffective(d);
  if (eff.deleted) return false;
  if (filters.state !== 'ALL' && eff.state !== filters.state) return false;
  if (filters.grade !== 'ALL' && eff.grade !== filters.grade) return false;
  if (filters.status !== 'ALL' && eff.status !== filters.status) return false;
  if (filters.vert !== 'ALL') {
    const v = eff.vertical.toLowerCase();
    const fv = filters.vert.toLowerCase();
    if (fv==='food' && !v.match(/food|meat|poultry|baked|cereal|fruit|vegetable|frozen|snack|candy|confect|flour|grain|sausage/)) return false;
    if (fv==='beverage' && !v.match(/beverage|brew|bottl|distill|wine|beer|coffee|coca|soda/)) return false;
    if (fv==='petfood' && !v.match(/pet|animal feed/)) return false;
    if (fv==='pharma' && !v.match(/pharma|medical|biotech|animal health|cosmetic|personal care/)) return false;
    if (fv==='packaging' && !v.match(/packag|container|carton|label/)) return false;
    if (fv==='aero' && !v.match(/aero|aircraft|aviation|defense|ammunition/)) return false;
    if (fv==='auto' && !v.match(/auto|transport|vehicle|tire|brake/)) return false;
    if (fv==='chemical' && !v.match(/chemical|lubricant|paint|adhesive|coating/)) return false;
    if (fv==='plastics' && !v.match(/plastic|rubber|polymer|extrusion|wire.*cable|cable/)) return false;
    if (fv==='building' && !v.match(/building|construct|hvac|lighting|furniture|hose/)) return false;
    if (fv==='industrial' && !v.match(/industrial|equipment|metal|steel|foundry|manufact/)) return false;
    if (fv==='electronics' && !v.match(/electr|electro|tech/)) return false;
    if (fv==='consumer' && !v.match(/consumer|cpg/)) return false;
    if (fv==='graphics' && !v.match(/graphic|print|text|distribut/)) return false;
    if (fv==='other' && !v.match(/other|unknown|postal|whol/) && !v.match(/^\d/)) return false;
  }
  if (filters.health !== 'ALL') {
    if (filters.health === 'none' && eff.healthGrade) return false;
    if (filters.health !== 'none' && eff.healthGrade !== filters.health) return false;
  }
  if (filters.upgrade !== 'ALL') {
    if (filters.upgrade === 'Y' && !eff.upgradeReady) return false;
    if (filters.upgrade === 'N' && eff.upgradeReady) return false;
  }
  if (filters.contacts !== 'ALL') {
    const cnt = eff.contacts ? eff.contacts.length : 0;
    if (filters.contacts === '0' && cnt !== 0) return false;
    if (filters.contacts === '1' && cnt !== 1) return false;
    if (filters.contacts === '2+' && cnt < 2) return false;
  }
  if (filters.opps !== 'ALL') {
    const hasOpps = eff.oppNote && eff.oppNote.length > 0;
    if (filters.opps === 'Y' && !hasOpps) return false;
    if (filters.opps === 'N' && hasOpps) return false;
  }
  if (filters.rev !== 'ALL') {
    const r = eff.rev012 || 0;
    if (filters.rev === '0' && r !== 0) return false;
    if (filters.rev === '1-10k' && (r <= 0 || r >= 10000)) return false;
    if (filters.rev === '10-50k' && (r < 10000 || r >= 50000)) return false;
    if (filters.rev === '50-100k' && (r < 50000 || r >= 100000)) return false;
    if (filters.rev === '100k+' && r < 100000) return false;
  }
  if (filters.starred !== 'ALL') {
    if (filters.starred === 'Y' && !eff.star) return false;
    if (filters.starred === 'N' && eff.star) return false;
  }
  if (filters.commented !== 'ALL') {
    const hasComments = (eff.note && eff.note.trim().length > 0) || (eff.callNotes && eff.callNotes.length > 0);
    if (filters.commented === 'Y' && !hasComments) return false;
    if (filters.commented === 'N' && hasComments) return false;
  }
  if (filters.search) {
    const s = filters.search.toLowerCase();
    if (!eff.company.toLowerCase().includes(s) && !eff.city.toLowerCase().includes(s) && !eff.vertical.toLowerCase().includes(s) && !eff.state.toLowerCase().includes(s) && !eff.address.toLowerCase().includes(s)) return false;
  }
  return true;
}

function getMarkerClass(d) {
  const eff = getEffective(d);
  if (eff.status === 'active') return 'status-active';
  if (eff.status === 'lapsed') return 'status-lapsed';
  if (eff.status === 'dormant') return 'status-dormant';
  return 'grade-' + eff.grade;
}

function buildPopup(d, idx) {
  const eff = getEffective(d, idx);
  const hasContacts = eff.contacts && eff.contacts.length > 0;
  const statusLabel = eff.status === 'active' ? 'Active Customer' : eff.status === 'lapsed' ? 'Lapsed Customer' : eff.status === 'dormant' ? 'Dormant Customer' : 'Prospect';
  const statusClass = eff.status === 'active' ? 'badge-active' : eff.status === 'lapsed' ? 'badge-lapsed' : eff.status === 'dormant' ? 'badge-dormant' : '';
  const gradeClass = 'badge-' + eff.grade.toLowerCase();
  const trendIcon = eff.trend === 'Rising' ? '<span style="color:var(--a-color)">&uarr; Growing</span>' : eff.trend === 'Declining' ? '<span style="color:#FF6B6B">&darr; Declining</span>' : '&mdash;';

  return \`<div class="popup-inner">
    <div class="popup-company">\${eff.star ? '&#x2B50; ' : ''}\${eff.company}</div>
    <div class="popup-row">&#x1F4CD; <strong>\${eff.address ? eff.address + ', ' : ''}\${eff.city}, \${eff.state}</strong></div>
    <div class="popup-badges">
      <span class="badge \${statusClass || 'badge-vert'}">\${statusLabel}</span>
      <span class="badge \${gradeClass}">OPP: \${eff.grade} (\${eff.score}pts)</span>
      \${eff.healthGrade ? \`<span class="badge badge-\${eff.healthGrade === 'A' ? 'a' : eff.healthGrade === 'B' ? 'b' : eff.healthGrade === 'C' ? 'c' : 'd'}">HEALTH: \${eff.healthGrade} (\${eff.healthScore}pts)</span>\` : ''}
      <span class="badge badge-vert">\${eff.vertical}</span>
    </div>
    <div class="popup-details">
      \${eff.rev012 ? \`<div class="popup-detail-row"><span>Rev 0-12mo</span><span class="val green">\${formatSales(eff.rev012)}</span></div>\` : ''}
      \${eff.rev1224 ? \`<div class="popup-detail-row"><span>Rev 12-24mo</span><span class="val">\${formatSales(eff.rev1224)}</span></div>\` : ''}
      \${eff.revPrvYTD ? \`<div class="popup-detail-row"><span>Prev YTD</span><span class="val">\${formatSales(eff.revPrvYTD)}</span></div>\` : ''}
      \${eff.revYOYDiff ? \`<div class="popup-detail-row"><span>YOY Change</span><span class="val \${eff.revYOYDiff > 0 ? 'green' : 'red'}">\${eff.revYOYDiff > 0 ? '+' : ''}\${formatSales(eff.revYOYDiff)}</span></div>\` : ''}
      <div class="popup-detail-row"><span>YOY Trend</span><span class="val">\${trendIcon}</span></div>
      \${eff.isContractCustomer ? '<div class="popup-detail-row"><span>Service Contract</span><span class="val green">Yes</span></div>' : ''}
      <div class="popup-detail-row"><span>Upgrade Ready</span><span class="val">\${eff.upgradeReady ? 'Yes' : 'No'}</span></div>
      <div class="popup-detail-row"><span>VJ Systems</span><span class="val">\${eff.prodLines}</span></div>
      \${eff.apRev012 ? \`<div class="popup-detail-row"><span>Equip Purchased 0-12mo</span><span class="val green">\${formatSales(eff.apRev012)}</span></div>\` : ''}
      \${eff.apRev1224 ? \`<div class="popup-detail-row"><span>Equip Purchased 12-24mo</span><span class="val">\${formatSales(eff.apRev1224)}</span></div>\` : ''}
      \${eff.vjEquip && Object.keys(eff.vjEquip).length > 0 ? Object.entries(eff.vjEquip).map(([tech, count]) => \`<div class="popup-detail-row" style="padding-left:12px"><span>\${tech}</span><span class="val">\${count}</span></div>\`).join('') : ''}
      <div class="popup-detail-row"><span>Comp. CIJ</span><span class="val">\${eff.compCIJ}</span></div>
      <div class="popup-detail-row"><span>Comp. Laser</span><span class="val">\${eff.compLaser}</span></div>
      \${eff.oppNote ? \`<div class="popup-detail-row"><span>Opportunities</span><span class="val">\${eff.oppNote}</span></div>\` : ''}
    </div>
    \${eff.note ? \`<div class="popup-note">\${eff.note}</div>\` : ''}
    <div class="popup-actions">
      \${hasContacts ? \`<div class="popup-btn primary" onclick="showContacts(\${idx})">&#x1F465; Contacts (\${eff.contacts.length})</div>\` : '<div style="font-size:10px;color:var(--text-dim);font-style:italic;margin-top:4px;font-family:IBM Plex Mono,monospace">No contacts on file</div>'}
      <div class="popup-btn" onclick="openEditModal(\${idx})">&#x270F;&#xFE0F; Edit</div>
    </div>
  </div>\`;
}

function render() {
  markers = [];
  plainGroup.clearLayers();

  // Include custom-added accounts from localStorage
  const edits = loadEdits();
  const customAccounts = [];
  Object.keys(edits).forEach(key => {
    const e = edits[key];
    if (e.isCustom && !e.deleted) {
      customAccounts.push(e.data);
    }
  });
  const allData = [...RAW_DATA, ...customAccounts];

  const filtered = allData.map((d,i) => ({...d, _idx:i})).filter(matchesFilters);

  // Counts
  let activeCount=0, lapsedCount=0, dormantCount=0, prospectCount=0;
  filtered.forEach(d => {
    const eff = getEffective(d);
    if (eff.status === 'active') activeCount++;
    else if (eff.status === 'lapsed') lapsedCount++;
    else if (eff.status === 'dormant') dormantCount++;
    else prospectCount++;
  });

  document.getElementById('result-count').textContent = filtered.length;
  document.getElementById('hdr-total').textContent = filtered.length;
  document.getElementById('hdr-active').textContent = activeCount;
  document.getElementById('hdr-lapsed').textContent = lapsedCount;
  document.getElementById('hdr-dormant').textContent = dormantCount;
  document.getElementById('hdr-prospect').textContent = prospectCount;

  // Sidebar list
  const listEl = document.getElementById('results-list');
  if (filtered.length === 0) {
    listEl.innerHTML = '<div style="padding:20px 16px;color:var(--text-dim);font-size:12px;text-align:center;">No accounts match filters</div>';
  } else {
    listEl.innerHTML = filtered.slice(0, 200).map(d => {
      const eff = getEffective(d);
      return \`<div class="result-item status-\${eff.status} grade-\${eff.grade} \${selectedIdx === d._idx ? 'selected' : ''}" data-idx="\${d._idx}" onclick="selectAccount(\${d._idx})">
        <div class="result-company">\${eff.star ? '<span class="result-star">&#x2B50;</span>' : ''}\${eff.ring ? '<span class="result-ring">&#x1F534;</span>' : ''}\${eff.company}</div>
        <div class="result-meta">
          <span class="result-grade \${eff.grade}">\${eff.grade}</span>
          <span>\${eff.city}, \${eff.state}</span>
          \${eff.status === 'active' ? '<span class="result-status-badge active">ACTIVE</span>' : ''}
          \${eff.status === 'lapsed' ? '<span class="result-status-badge lapsed">LAPSED</span>' : eff.status === 'dormant' ? '<span class="result-status-badge dormant">DORMANT</span>' : ''}
        </div>
      </div>\`;
    }).join('');
    if (filtered.length > 200) {
      listEl.innerHTML += \`<div style="padding:12px 16px;color:var(--text-dim);font-size:11px;text-align:center;font-family:'IBM Plex Mono',monospace;">+\${filtered.length - 200} more — use filters to narrow</div>\`;
    }
  }

  // Render individual markers (no clustering)
  filtered.forEach(d => {
    const eff = getEffective(d);
    const cls = getMarkerClass(d);
    const starHTML = eff.star ? '<span class="star-overlay">&#x2B50;</span>' : '';
    const ringHTML = eff.ring ? '<span class="ring-overlay"></span>' : '';
    const icon = L.divIcon({
      className: '',
      html: \`<div class="map-marker \${cls} \${selectedIdx === d._idx ? 'selected' : ''}" data-idx="\${d._idx}">\${starHTML}\${ringHTML}</div>\`,
      iconSize: [14,14], iconAnchor: [7,7]
    });
    const marker = L.marker([d.lat, d.lng], { icon, zIndexOffset: d.status==='active'?200:d.status==='lapsed'?150:d.status==='dormant'?100:d.grade==='A'?50:0 })
      .bindPopup(buildPopup(d, d._idx), { maxWidth: 360 });
    marker.on('click', () => selectAccount(d._idx));
    plainGroup.addLayer(marker);
    markers.push({ marker, idx: d._idx });
  });
}

function selectAccount(idx) {
  selectedIdx = idx;
  const allData = [...RAW_DATA, ...(function(){ const edits=loadEdits(); const ca=[]; Object.keys(edits).forEach(k=>{const e=edits[k];if(e.isCustom&&!e.deleted)ca.push(e.data)}); return ca; })()];
  const d = allData[idx];
  if (!d) return;
  map.setView([d.lat, d.lng], Math.max(map.getZoom(), 10));
  const m = markers.find(m => m.idx === idx);
  if (m) m.marker.openPopup();
  document.querySelectorAll('.result-item').forEach(el => {
    el.classList.toggle('selected', parseInt(el.dataset.idx) === idx);
  });
  const el = document.querySelector(\`.result-item[data-idx="\${idx}"]\`);
  if (el) el.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

// Filter handlers
function setupPillGroup(groupId, filterKey) {
  document.getElementById(groupId).querySelectorAll('.pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.getElementById(groupId).querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      filters[filterKey] = pill.dataset.val;
      selectedIdx = null;
      render();
    });
  });
}
setupPillGroup('state-filter', 'state');
setupPillGroup('grade-filter', 'grade');
setupPillGroup('status-filter', 'status');
setupPillGroup('vert-filter', 'vert');
setupPillGroup('health-filter', 'health');
setupPillGroup('upgrade-filter', 'upgrade');
setupPillGroup('contacts-filter', 'contacts');
setupPillGroup('opps-filter', 'opps');
setupPillGroup('rev-filter', 'rev');
setupPillGroup('starred-filter', 'starred');
setupPillGroup('commented-filter', 'commented');

document.getElementById('search').addEventListener('input', e => {
  filters.search = e.target.value.trim();
  selectedIdx = null;
  render();
});

document.getElementById('clear-filters').addEventListener('click', () => {
  filters = { state:'ALL', grade:'ALL', status:'ALL', vert:'ALL', health:'ALL', upgrade:'ALL', contacts:'ALL', opps:'ALL', rev:'ALL', starred:'ALL', commented:'ALL', search:'' };
  document.getElementById('search').value = '';
  document.querySelectorAll('.pill').forEach(p => p.classList.toggle('active', p.dataset.val === 'ALL'));
  selectedIdx = null;
  render();
});

// Contacts panel
window.showContacts = function(idx) {
  const allData = [...RAW_DATA, ...(function(){ const edits=loadEdits(); const ca=[]; Object.keys(edits).forEach(k=>{const e=edits[k];if(e.isCustom&&!e.deleted)ca.push(e.data)}); return ca; })()];
  const d = allData[idx];
  const eff = getEffective(d, idx);
  document.getElementById('cp-company-name').textContent = eff.company;
  const listEl = document.getElementById('cp-contact-list');

  if (!eff.contacts || eff.contacts.length === 0) {
    listEl.innerHTML = '<div class="cp-empty">No contacts on file for this account</div>';
  } else {
    listEl.innerHTML = eff.contacts.map(c => {
      const hasPhone = c.phone && c.phone.trim() !== '';
      const hasMobile = c.mobile && c.mobile.trim() !== '';
      const hasEmail = c.email && c.email.trim() !== '';
      const clean = n => n.replace(/[^0-9+]/g, '');
      const isEquip = c.title === 'Equipment Contact';
      const fitClass = isEquip ? 'equip' : (c.fit === 'A' || c.fit === 'B') ? 'fit-' + c.fit : 'fit-other';
      return \`<div class="cp-contact \${fitClass}">
        <div class="cp-name">
          \${c.name}
          \${c.dm ? '<span class="cp-dm-badge">&#x2605; DM</span>' : ''}
          \${isEquip ? '<span class="cp-equip-badge">EQUIP</span>' : ''}
        </div>
        <div class="cp-title">\${c.title || 'Title unknown'}</div>
        \${hasPhone ? \`<div class="cp-phone-row">&#x1F4DE; <a class="cp-phone" href="tel:+1\${clean(c.phone)}">\${c.phone}</a></div>\` : ''}
        \${hasMobile ? \`<div class="cp-phone-row">&#x1F4F1; <a class="cp-phone" href="tel:+1\${clean(c.mobile)}">\${c.mobile}</a></div>\` : ''}
        \${hasEmail ? \`<div class="cp-phone-row">&#x2709; <a class="cp-email" href="mailto:\${c.email}">\${c.email}</a></div>\` : ''}
        \${!hasPhone && !hasMobile && !hasEmail ? '<div class="cp-nophone">No contact info on file</div>' : ''}
      </div>\`;
    }).join('');
  }
  document.getElementById('contacts-panel').classList.add('open');
  map.closePopup();
};

window.closeContacts = function() {
  document.getElementById('contacts-panel').classList.remove('open');
};

map.on('click', function() { closeContacts(); });

// ── Edit Modal ──
let editingIdx = null;

window.openEditModal = function(idx) {
  editingIdx = idx;
  const allData = [...RAW_DATA, ...(function(){ const edits=loadEdits(); const ca=[]; Object.keys(edits).forEach(k=>{const e=edits[k];if(e.isCustom&&!e.deleted)ca.push(e.data)}); return ca; })()];
  const d = allData[idx];
  const eff = getEffective(d, idx);
  document.getElementById('edit-modal-title').textContent = 'Edit: ' + eff.company;
  document.getElementById('edit-star').value = eff.star ? '1' : '0';
  document.getElementById('edit-ring').value = eff.ring ? '1' : '0';
  document.getElementById('edit-grade').value = '';
  document.getElementById('edit-status').value = '';
  document.getElementById('edit-note').value = eff.note || '';

  // Check if there's an existing override
  const edits = loadEdits();
  const key = getEditKey(d);
  const e = edits[key] || {};
  if (e.grade) document.getElementById('edit-grade').value = e.grade;
  if (e.status) document.getElementById('edit-status').value = e.status;

  document.getElementById('edit-delete-btn').onclick = function() {
    if (confirm('Delete ' + eff.company + ' from the map?')) {
      const edits = loadEdits();
      const key = getEditKey(d);
      edits[key] = { ...(edits[key]||{}), deleted: true };
      saveEdits(edits);
      closeEditModal();
      map.closePopup();
      render();
    }
  };

  document.getElementById('edit-modal').classList.add('open');
  map.closePopup();
};

window.saveEdit = function() {
  const allData = [...RAW_DATA, ...(function(){ const edits=loadEdits(); const ca=[]; Object.keys(edits).forEach(k=>{const e=edits[k];if(e.isCustom&&!e.deleted)ca.push(e.data)}); return ca; })()];
  const d = allData[editingIdx];
  const edits = loadEdits();
  const key = getEditKey(d);
  const existing = edits[key] || {};

  const star = document.getElementById('edit-star').value === '1';
  const ring = document.getElementById('edit-ring').value === '1';
  const grade = document.getElementById('edit-grade').value || null;
  const status = document.getElementById('edit-status').value || null;
  const note = document.getElementById('edit-note').value.trim();

  edits[key] = { ...existing, star, ring, grade, status, note };
  saveEdits(edits);
  closeEditModal();
  render();
};

window.closeEditModal = function() {
  document.getElementById('edit-modal').classList.remove('open');
  editingIdx = null;
};

// ── Add Account ──
let addingPin = false;

window.openAddAccount = function() {
  document.getElementById('add-name').value = '';
  document.getElementById('add-address').value = '';
  document.getElementById('add-city').value = '';
  document.getElementById('add-state').value = 'KS';
  document.getElementById('add-grade').value = 'C';
  document.getElementById('add-note').value = '';
  document.getElementById('add-lat').value = '';
  document.getElementById('add-lng').value = '';
  document.getElementById('add-modal').classList.add('open');
  addingPin = true;
};

// Allow clicking map to set coords when adding
map.on('click', function(e) {
  if (addingPin) {
    document.getElementById('add-lat').value = e.latlng.lat.toFixed(4);
    document.getElementById('add-lng').value = e.latlng.lng.toFixed(4);
  }
});

window.saveNewAccount = function() {
  const name = document.getElementById('add-name').value.trim();
  const addr = document.getElementById('add-address').value.trim();
  const city = document.getElementById('add-city').value.trim();
  const state = document.getElementById('add-state').value;
  const grade = document.getElementById('add-grade').value;
  const note = document.getElementById('add-note').value.trim();
  const lat = parseFloat(document.getElementById('add-lat').value);
  const lng = parseFloat(document.getElementById('add-lng').value);

  if (!name) { alert('Please enter a company name'); return; }
  if (!lat || !lng || isNaN(lat) || isNaN(lng)) { alert('Please set coordinates (click on map or enter lat/lng)'); return; }

  const newAccount = {
    company: name, address: addr, city: city, state: state,
    vertical: 'Other / Unknown', lat: lat, lng: lng,
    status: 'prospect', grade: grade, score: 0,
    rev012: 0, rev1224: 0, trend: 'unknown',
    upgradeReady: false, compCIJ: 0, compLaser: 0,
    prodLines: 0, oppNote: '', contacts: []
  };

  const edits = loadEdits();
  const key = getEditKey(newAccount);
  edits[key] = { isCustom: true, data: newAccount, note: note, star: false, ring: false };
  saveEdits(edits);
  closeAddModal();
  render();
};

window.closeAddModal = function() {
  document.getElementById('add-modal').classList.remove('open');
  addingPin = false;
};

// ── Rules Panel ──
window.openRulesPanel = function() {
  document.getElementById('rules-panel').classList.add('open');
};
window.closeRulesPanel = function() {
  document.getElementById('rules-panel').classList.remove('open');
};

// ── Export Edits ──
window.exportEdits = function() {
  const edits = loadEdits();
  const blob = new Blob([JSON.stringify(edits, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'vj-territory-edits-' + new Date().toISOString().slice(0,10) + '.json';
  a.click();
  URL.revokeObjectURL(url);
};

// Initial render
render();
}); // end DOMContentLoaded
<\/script>
</body>
</html>`;

// ── MOBILE HTML ──
const mobileHTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no"/>
<title>VJ Territory Intel — Mobile</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"><\/script>
<style>
:root {
  --bg: #0a0e14;
  --surface: #111820;
  --surface2: #161e28;
  --border: #1e2d3d;
  --border2: #243344;
  --accent: #00d4ff;
  --text: #c8d8e8;
  --text-dim: #5a7a94;
  --text-bright: #e8f4ff;
  --active-color: #FFD700;
  --active-glow: rgba(255,215,0,0.4);
  --lapsed-color: #B71C1C;
  --lapsed-glow: rgba(183,28,28,0.5);
  --dormant-color: #F48FB1;
  --dormant-glow: rgba(244,143,177,0.4);
  --a-color: #2ECC71;
  --a-glow: rgba(46,204,113,0.35);
  --b-color: #3498DB;
  --b-glow: rgba(52,152,219,0.35);
  --c-color: #FF8C00;
  --c-glow: rgba(255,140,0,0.35);
  --d-color: #95A5A6;
  --d-glow: rgba(149,165,166,0.35);
}
* { margin:0; padding:0; box-sizing:border-box; }
html, body { height:100%; width:100%; overflow:hidden; background:var(--bg); color:var(--text); font-family:'IBM Plex Sans',sans-serif; }

#map { position:fixed; top:0; left:0; right:0; bottom:0; z-index:1; }

/* HEADER */
#header {
  position:fixed; top:0; left:0; right:0; z-index:500;
  background:rgba(10,14,20,0.95); border-bottom:1px solid var(--border);
  backdrop-filter:blur(8px); padding:8px 12px 6px;
}
#header-row1 { display:flex; align-items:center; justify-content:space-between; margin-bottom:6px; gap:8px; }
.logo { font-family:'IBM Plex Mono',monospace; font-size:11px; font-weight:700; letter-spacing:2px; color:var(--text-bright); text-transform:uppercase; display:flex; align-items:center; gap:6px; white-space:nowrap; flex-shrink:0; }
.logo-dot { width:6px; height:6px; background:var(--accent); border-radius:50%; box-shadow:0 0 8px var(--accent); animation:pulse 2s infinite; }
@keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:0.5;transform:scale(0.7)} }
.hamburger { font-size:18px; cursor:pointer; color:var(--text-dim); padding:4px 8px; border:1px solid var(--border2); border-radius:4px; }
.hamburger:active { background:var(--surface2); }
#search-wrap { position:relative; flex:1; }
#search-wrap input { width:100%; background:var(--surface2); border:1px solid var(--border2); border-radius:4px; color:var(--text-bright); padding:7px 10px 7px 28px; font-family:'IBM Plex Sans',sans-serif; font-size:13px; outline:none; }
#search-wrap input:focus { border-color:var(--accent); }
#search-wrap input::placeholder { color:var(--text-dim); }
.search-icon { position:absolute; left:8px; top:50%; transform:translateY(-50%); color:var(--text-dim); font-size:13px; pointer-events:none; }
#stats-row { display:flex; gap:6px; overflow-x:auto; padding-bottom:2px; scrollbar-width:none; }
#stats-row::-webkit-scrollbar { display:none; }
.stat-chip { font-family:'IBM Plex Mono',monospace; font-size:10px; white-space:nowrap; background:var(--surface2); border:1px solid var(--border2); border-radius:3px; padding:3px 8px; display:flex; align-items:center; gap:4px; }
.stat-chip .val { font-size:13px; font-weight:600; }

.leaflet-top { margin-top:82px; }

/* SIDEBAR (mobile) */
#sidebar {
  position:fixed; top:0; left:0; bottom:0; width:280px; z-index:600;
  background:var(--surface); border-right:1px solid var(--border2);
  transform:translateX(-100%); transition:transform 0.3s ease;
  display:flex; flex-direction:column; overflow:hidden;
}
#sidebar.open { transform:translateX(0); }
#sidebar-overlay {
  position:fixed; top:0; left:0; right:0; bottom:0; z-index:599;
  background:rgba(0,0,0,0.5); display:none;
}
#sidebar-overlay.open { display:block; }
#sidebar-header { padding:14px 16px; border-bottom:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; }
#sidebar-header h3 { font-family:'IBM Plex Mono',monospace; font-size:11px; text-transform:uppercase; letter-spacing:2px; color:var(--text-dim); }
#sidebar-close { color:var(--text-dim); font-size:20px; cursor:pointer; }
#sidebar-body { flex:1; overflow-y:auto; padding:14px 16px; }
.filter-section { margin-bottom:14px; }
.filter-lbl { font-family:'IBM Plex Mono',monospace; font-size:9px; text-transform:uppercase; letter-spacing:1.5px; color:var(--text-dim); margin-bottom:6px; }
.filter-pills { display:flex; flex-wrap:wrap; gap:5px; }
.pill { padding:4px 10px; border-radius:3px; font-size:11px; font-family:'IBM Plex Mono',monospace; cursor:pointer; border:1px solid var(--border2); background:var(--bg); color:var(--text-dim); transition:all 0.15s; user-select:none; }
.pill:active { opacity:0.7; }
.pill.active { background:var(--accent); border-color:var(--accent); color:#000; font-weight:600; }
.pill.grade-a.active { background:var(--a-color); border-color:var(--a-color); }
.pill.grade-b.active { background:var(--b-color); border-color:var(--b-color); }
.pill.grade-c.active { background:var(--c-color); border-color:var(--c-color); }
.pill.status-active.active { background:var(--active-color); border-color:var(--active-color); }
.pill.status-lapsed.active { background:var(--lapsed-color); border-color:var(--lapsed-color); }
.pill.status-dormant.active { background:var(--dormant-color); border-color:var(--dormant-color); }
#sidebar-actions { padding:14px 16px; border-top:1px solid var(--border); display:flex; flex-direction:column; gap:8px; }
.sb-btn { width:100%; padding:8px; border:1px solid var(--border2); border-radius:4px; background:none; color:var(--text-dim); font-family:'IBM Plex Mono',monospace; font-size:10px; letter-spacing:1px; text-transform:uppercase; cursor:pointer; text-align:center; }
.sb-btn:active { background:var(--surface2); }

#legend-row { display:flex; gap:10px; flex-wrap:wrap; padding-top:10px; margin-top:10px; border-top:1px solid var(--border); }
.leg-item { display:flex; align-items:center; gap:5px; font-size:9px; color:var(--text-dim); }
.leg-dot { width:8px; height:8px; border-radius:50%; }

/* Popups */
.leaflet-popup-content-wrapper { background:var(--surface)!important; border:1px solid var(--border2)!important; border-radius:6px!important; box-shadow:0 8px 32px rgba(0,0,0,0.6)!important; padding:0!important; }
.leaflet-popup-tip { background:var(--surface)!important; }
.leaflet-popup-close-button { color:var(--text-dim)!important; }
.leaflet-popup-content { margin:0!important; padding:0!important; width:auto!important; }
.leaflet-control-attribution { display:none; }
.popup-inner { padding:12px 14px; min-width:220px; max-width:300px; }
.popup-company { font-family:'IBM Plex Mono',monospace; font-size:14px; font-weight:600; color:var(--text-bright); margin-bottom:6px; }
.popup-row { font-size:11px; color:var(--text-dim); margin-bottom:3px; font-family:'IBM Plex Mono',monospace; }
.popup-badges { display:flex; gap:5px; margin-top:8px; flex-wrap:wrap; }
.badge { padding:2px 7px; border-radius:3px; font-size:10px; font-family:'IBM Plex Mono',monospace; font-weight:600; }
.badge-a { background:rgba(46,204,113,0.12); color:var(--a-color); border:1px solid rgba(46,204,113,0.3); }
.badge-b { background:rgba(52,152,219,0.12); color:var(--b-color); border:1px solid rgba(52,152,219,0.3); }
.badge-c { background:rgba(255,140,0,0.12); color:var(--c-color); border:1px solid rgba(255,140,0,0.3); }
.badge-d { background:rgba(149,165,166,0.12); color:var(--d-color); border:1px solid rgba(149,165,166,0.3); }
.badge-active { background:rgba(255,215,0,0.12); color:var(--active-color); border:1px solid rgba(255,215,0,0.3); }
.badge-lapsed { background:rgba(183,28,28,0.12); color:var(--lapsed-color); border:1px solid rgba(183,28,28,0.3); }
.badge-dormant { background:rgba(244,143,177,0.12); color:var(--dormant-color); border:1px solid rgba(244,143,177,0.3); }
.badge-vert { background:rgba(0,212,255,0.08); color:var(--accent); border:1px solid rgba(0,212,255,0.2); }
.popup-details { margin-top:8px; padding-top:8px; border-top:1px solid var(--border); font-family:'IBM Plex Mono',monospace; font-size:10px; }
.popup-detail-row { display:flex; justify-content:space-between; margin-bottom:3px; color:var(--text-dim); }
.popup-detail-row .val { color:var(--text); font-weight:500; }
.popup-detail-row .val.green { color:var(--a-color); }
.popup-detail-row .val.red { color:#FF6B6B; }
.popup-note { margin-top:6px; padding:6px 8px; background:rgba(0,212,255,0.05); border:1px solid var(--border); border-radius:3px; font-size:10px; color:var(--text); font-family:'IBM Plex Mono',monospace; }
.popup-actions { margin-top:10px; display:flex; gap:6px; flex-wrap:wrap; }
.popup-btn { display:inline-flex; align-items:center; gap:5px; padding:6px 10px; border-radius:4px; cursor:pointer; font-family:'IBM Plex Mono',monospace; font-size:10px; font-weight:600; border:1px solid var(--border2); color:var(--text-dim); background:var(--bg); transition:all 0.15s; }
.popup-btn.primary { border-color:var(--accent); color:var(--accent); background:rgba(0,212,255,0.08); }

/* Contacts panel */
#contacts-panel {
  position:fixed; top:0; bottom:0; left:0; right:0; z-index:800;
  background:var(--surface); display:flex; flex-direction:column;
  transform:translateX(100%); transition:transform 0.25s ease;
}
#contacts-panel.open { transform:translateX(0); }
.cp-header { padding:14px 16px; border-bottom:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; }
.cp-company { font-family:'IBM Plex Mono',monospace; font-size:13px; font-weight:600; color:var(--text-bright); flex:1; padding-right:10px; }
.cp-close { color:var(--text-dim); font-size:22px; cursor:pointer; }
.cp-list { flex:1; overflow-y:auto; padding:6px 0; }
.cp-contact { padding:10px 16px; border-bottom:1px solid var(--border); position:relative; }
.cp-contact::before { content:''; position:absolute; left:0; top:0; bottom:0; width:3px; }
.cp-contact.equip::before { background:var(--accent); }
.cp-contact.fit-A::before { background:var(--a-color); }
.cp-contact.fit-B::before { background:var(--b-color); }
.cp-contact.fit-other::before { background:var(--text-dim); opacity:0.4; }
.cp-name { font-size:13px; font-weight:600; color:var(--text-bright); margin-bottom:2px; display:flex; align-items:center; gap:7px; }
.cp-dm { font-family:'IBM Plex Mono',monospace; font-size:8px; font-weight:700; padding:1px 5px; border-radius:2px; background:rgba(255,215,0,0.12); color:var(--active-color); border:1px solid rgba(255,215,0,0.3); }
.cp-equip { font-family:'IBM Plex Mono',monospace; font-size:8px; font-weight:700; padding:1px 5px; border-radius:2px; background:rgba(0,212,255,0.12); color:var(--accent); border:1px solid rgba(0,212,255,0.3); }
.cp-title { font-size:11px; color:var(--text); margin-bottom:4px; }
.cp-phone-row { font-size:12px; margin-top:2px; }
.cp-phone-row a { font-family:'IBM Plex Mono',monospace; color:var(--accent); text-decoration:none; }
.cp-email-row { font-size:11px; margin-top:2px; }
.cp-email-row a { font-family:'IBM Plex Mono',monospace; color:var(--accent); text-decoration:none; font-size:10px; }
.cp-nophone { font-size:10px; color:var(--text-dim); font-style:italic; margin-top:3px; font-family:'IBM Plex Mono',monospace; }
.cp-empty { padding:24px 16px; color:var(--text-dim); font-family:'IBM Plex Mono',monospace; font-size:11px; text-align:center; }

/* Edit modal */
.edit-overlay { position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.6); z-index:900; display:none; align-items:center; justify-content:center; padding:16px; }
.edit-overlay.open { display:flex; }
.edit-modal { background:var(--surface); border:1px solid var(--border2); border-radius:8px; padding:18px; width:100%; max-width:360px; box-shadow:0 16px 48px rgba(0,0,0,0.5); }
.edit-modal h3 { font-family:'IBM Plex Mono',monospace; font-size:13px; color:var(--text-bright); margin-bottom:12px; }
.edit-row { display:flex; align-items:center; gap:8px; margin-bottom:8px; }
.edit-row label { font-family:'IBM Plex Mono',monospace; font-size:10px; color:var(--text-dim); min-width:60px; }
.edit-row select, .edit-row input, .edit-row textarea { background:var(--bg); border:1px solid var(--border2); border-radius:4px; color:var(--text-bright); padding:6px 10px; font-family:'IBM Plex Sans',sans-serif; font-size:12px; outline:none; flex:1; }
.edit-row textarea { resize:vertical; min-height:50px; }
.edit-btns { display:flex; gap:8px; margin-top:12px; justify-content:flex-end; }
.edit-btn { padding:7px 14px; border-radius:4px; font-family:'IBM Plex Mono',monospace; font-size:11px; cursor:pointer; border:1px solid var(--border2); background:var(--bg); color:var(--text-dim); }
.edit-btn.save { border-color:var(--a-color); color:var(--a-color); }
.edit-btn.danger { border-color:#FF6B6B; color:#FF6B6B; }

/* Map markers */
.map-marker { width:12px; height:12px; border-radius:50%; border:2px solid rgba(255,255,255,0.25); cursor:pointer; transition:transform 0.15s; position:relative; }
.map-marker:hover { transform:scale(1.6); }
.map-marker.status-active { background:var(--active-color); box-shadow:0 0 10px var(--active-glow); }
.map-marker.status-lapsed { background:var(--lapsed-color); box-shadow:0 0 12px var(--lapsed-glow); }
.map-marker.status-dormant { background:var(--dormant-color); box-shadow:0 0 12px var(--dormant-glow); }
.map-marker.grade-A { background:var(--a-color); box-shadow:0 0 8px var(--a-glow); }
.map-marker.grade-B { background:var(--b-color); box-shadow:0 0 8px var(--b-glow); }
.map-marker.grade-C { background:var(--c-color); box-shadow:0 0 8px var(--c-glow); }
.map-marker.grade-D { background:var(--d-color); box-shadow:0 0 6px var(--d-glow); }
.map-marker .star-overlay { position:absolute; top:-7px; right:-7px; font-size:8px; pointer-events:none; }
.map-marker .ring-overlay { position:absolute; top:-2px; left:-2px; right:-2px; bottom:-2px; border-radius:50%; border:2px solid #FF4444; pointer-events:none; animation:ringPulse 1.5s infinite; }
@keyframes ringPulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
.my-cluster { border-radius:50%; background:rgba(17,24,32,0.92); border:2px solid #243344; display:flex; align-items:center; justify-content:center; font-family:'IBM Plex Mono',monospace; font-weight:600; color:#e8f4ff; box-shadow:0 2px 10px rgba(0,0,0,0.5); cursor:pointer; }
</style>
</head>
<body>

<div id="header">
  <div id="header-row1">
    <div class="hamburger" onclick="toggleSidebar()">&#x2630;</div>
    <div class="logo"><div class="logo-dot"></div>VJ TERRITORY</div>
    <div id="search-wrap">
      <span class="search-icon">&#x2315;</span>
      <input type="text" id="search" placeholder="Search...">
    </div>
  </div>
  <div id="stats-row">
    <div class="stat-chip"><span class="val" style="color:var(--text-bright)" id="s-total">0</span> Accounts</div>
    <div class="stat-chip"><span class="val" style="color:var(--active-color)" id="s-active">0</span> Active</div>
    <div class="stat-chip"><span class="val" style="color:var(--lapsed-color)" id="s-lapsed">0</span> Lapsed</div>
    <div class="stat-chip"><span class="val" style="color:var(--dormant-color)" id="s-dormant">0</span> Dormant</div>
    <div class="stat-chip"><span class="val" style="color:var(--accent)" id="s-prospect">0</span> Prospect</div>
  </div>
</div>

<div id="map"></div>

<!-- Sidebar -->
<div id="sidebar-overlay" onclick="toggleSidebar()"></div>
<div id="sidebar">
  <div id="sidebar-header">
    <h3>Filters</h3>
    <div id="sidebar-close" onclick="toggleSidebar()">&#x2715;</div>
  </div>
  <div id="sidebar-body">
    <div class="filter-section">
      <div class="filter-lbl">State</div>
      <div class="filter-pills" id="f-state">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="MO">MO</div>
        <div class="pill" data-val="KS">KS</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Grade</div>
      <div class="filter-pills" id="f-grade">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill grade-a" data-val="A">A</div>
        <div class="pill grade-b" data-val="B">B</div>
        <div class="pill grade-c" data-val="C">C</div>
        <div class="pill" data-val="D">D</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Status</div>
      <div class="filter-pills" id="f-status">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill status-active" data-val="active">Active</div>
        <div class="pill status-lapsed" data-val="lapsed">Lapsed</div>
        <div class="pill status-dormant" data-val="dormant">Dormant</div>
        <div class="pill" data-val="prospect">Prospect</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Vertical / Industry</div>
      <div class="filter-pills" id="f-vert">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="Food">Food</div>
        <div class="pill" data-val="Beverage">Bev</div>
        <div class="pill" data-val="PetFood">Pet</div>
        <div class="pill" data-val="Pharma">Pharma</div>
        <div class="pill" data-val="Packaging">Pkg</div>
        <div class="pill" data-val="Aero">Aero</div>
        <div class="pill" data-val="Auto">Auto</div>
        <div class="pill" data-val="Chemical">Chem</div>
        <div class="pill" data-val="Plastics">Plastic</div>
        <div class="pill" data-val="Building">Build</div>
        <div class="pill" data-val="Industrial">Ind</div>
        <div class="pill" data-val="Electronics">Elec</div>
        <div class="pill" data-val="Consumer">CPG</div>
        <div class="pill" data-val="Graphics">Graph</div>
        <div class="pill" data-val="Other">Other</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Health Grade</div>
      <div class="filter-pills" id="f-health">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill grade-a" data-val="A">A</div>
        <div class="pill grade-b" data-val="B">B</div>
        <div class="pill grade-c" data-val="C">C</div>
        <div class="pill" data-val="D">D</div>
        <div class="pill" data-val="none">N/A</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Upgrade Ready</div>
      <div class="filter-pills" id="f-upgrade">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="Y">Yes</div>
        <div class="pill" data-val="N">No</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Contacts</div>
      <div class="filter-pills" id="f-contacts">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="0">0</div>
        <div class="pill" data-val="1">1</div>
        <div class="pill" data-val="2+">2+</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Open Opps</div>
      <div class="filter-pills" id="f-opps">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="Y">Has Opps</div>
        <div class="pill" data-val="N">None</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Revenue (12mo)</div>
      <div class="filter-pills" id="f-rev">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="0">$0</div>
        <div class="pill" data-val="1-10k">&lt;$10K</div>
        <div class="pill" data-val="10-50k">$10-50K</div>
        <div class="pill" data-val="50-100k">$50-100K</div>
        <div class="pill" data-val="100k+">$100K+</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Starred</div>
      <div class="filter-pills" id="f-starred">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="Y">&#x2B50; Starred</div>
        <div class="pill" data-val="N">Not Starred</div>
      </div>
    </div>
    <div class="filter-section">
      <div class="filter-lbl">Comments</div>
      <div class="filter-pills" id="f-commented">
        <div class="pill active" data-val="ALL">All</div>
        <div class="pill" data-val="Y">Has Comments</div>
        <div class="pill" data-val="N">No Comments</div>
      </div>
    </div>
    <div id="legend-row">
      <div class="leg-item"><div class="leg-dot" style="background:var(--active-color)"></div> Active</div>
      <div class="leg-item"><div class="leg-dot" style="background:var(--lapsed-color)"></div> Lapsed</div>
      <div class="leg-item"><div class="leg-dot" style="background:var(--dormant-color)"></div> Dormant</div>
      <div class="leg-item"><div class="leg-dot" style="background:var(--a-color)"></div> A-Grade</div>
      <div class="leg-item"><div class="leg-dot" style="background:var(--b-color)"></div> B-Grade</div>
      <div class="leg-item"><div class="leg-dot" style="background:var(--c-color)"></div> C-Grade</div>
      <div class="leg-item"><div class="leg-dot" style="background:var(--d-color)"></div> D-Grade</div>
    </div>
  </div>
  <div id="sidebar-actions">
    <div class="sb-btn" onclick="openAddAccount()">+ Add Account</div>
    <div class="sb-btn" onclick="exportEdits()">Export Edits</div>
    <div class="sb-btn" onclick="resetFilters()">Reset Filters</div>
    <div class="sb-btn" onclick="openRulesPanel()">Rules &amp; Scoring</div>
  </div>
</div>

<!-- Contacts panel -->
<div id="contacts-panel">
  <div class="cp-header">
    <div class="cp-company" id="cp-name"></div>
    <div class="cp-close" onclick="closeContacts()">&#x2715;</div>
  </div>
  <div class="cp-list" id="cp-list"></div>
</div>

<!-- Edit modal -->
<div class="edit-overlay" id="edit-modal">
  <div class="edit-modal">
    <h3 id="edit-title">Edit Account</h3>
    <div class="edit-row"><label>Star</label><select id="edit-star"><option value="0">No</option><option value="1">&#x2B50; Yes</option></select></div>
    <div class="edit-row"><label>Trial</label><select id="edit-ring"><option value="0">No</option><option value="1">&#x1F534; Active</option></select></div>
    <div class="edit-row"><label>Grade</label><select id="edit-grade"><option value="">Auto</option><option value="A">A</option><option value="B">B</option><option value="C">C</option><option value="D">D</option></select></div>
    <div class="edit-row"><label>Status</label><select id="edit-status"><option value="">Auto</option><option value="active">Active</option><option value="lapsed">Lapsed</option><option value="dormant">Dormant</option><option value="prospect">Prospect</option></select></div>
    <div class="edit-row"><label>Note</label><textarea id="edit-note" placeholder="Note..."></textarea></div>
    <div class="edit-btns">
      <div class="edit-btn danger" id="edit-delete">Delete</div>
      <div class="edit-btn" onclick="closeEditModal()">Cancel</div>
      <div class="edit-btn save" onclick="saveEdit()">Save</div>
    </div>
  </div>
</div>

<!-- Add modal -->
<div class="edit-overlay" id="add-modal">
  <div class="edit-modal">
    <h3>Add Account</h3>
    <div class="edit-row"><label>Name</label><input id="add-name" type="text" placeholder="Company"></div>
    <div class="edit-row"><label>Address</label><input id="add-address" type="text"></div>
    <div class="edit-row"><label>City</label><input id="add-city" type="text"></div>
    <div class="edit-row"><label>State</label><select id="add-state"><option value="KS">KS</option><option value="MO">MO</option></select></div>
    <div class="edit-row"><label>Grade</label><select id="add-grade"><option value="C">C</option><option value="A">A</option><option value="B">B</option><option value="D">D</option></select></div>
    <div class="edit-row"><label>Note</label><textarea id="add-note"></textarea></div>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--text-dim);margin-bottom:8px;">Tap map to set pin, or enter coords:</div>
    <div class="edit-row"><label>Lat</label><input id="add-lat" type="number" step="0.0001"></div>
    <div class="edit-row"><label>Lng</label><input id="add-lng" type="number" step="0.0001"></div>
    <div class="edit-btns">
      <div class="edit-btn" onclick="closeAddModal()">Cancel</div>
      <div class="edit-btn save" onclick="saveNewAccount()">Add</div>
    </div>
  </div>
</div>

<!-- Rules Panel -->
<div class="edit-overlay" id="rules-panel">
  <div class="edit-modal" style="max-height:85vh;overflow-y:auto">
    <h3>Rules &amp; Scoring</h3>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text);line-height:1.6">

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:1.5px;margin:10px 0 5px">Statuses</div>
      <div style="margin-bottom:3px"><span style="color:var(--active-color)">&#x25CF;</span> <strong>Active</strong> &mdash; VJ revenue last 12mo</div>
      <div style="margin-bottom:3px"><span style="color:var(--lapsed-color)">&#x25CF;</span> <strong>Lapsed</strong> &mdash; CRM says Active Customer, but zero revenue last 12mo</div>
      <div style="margin-bottom:3px"><span style="color:var(--dormant-color)">&#x25CF;</span> <strong>Dormant</strong> &mdash; Has VJ equipment but no orders last 12mo. CRM types as Prospect.</div>
      <div style="margin-bottom:3px"><span style="color:var(--a-color)">&#x25CF;</span><span style="color:var(--b-color)">&#x25CF;</span><span style="color:var(--c-color)">&#x25CF;</span><span style="color:var(--d-color)">&#x25CF;</span> <strong>Prospect</strong> &mdash; No recent VJ revenue. Color = opp grade.</div>

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:1.5px;margin:12px 0 5px">Opportunity Grade</div>
      <div style="margin-bottom:4px"><em>How likely do they need marking/coding?</em></div>
      <div style="padding-left:8px;margin-bottom:2px"><strong>Vertical</strong> (30) &middot; <strong>Plant size</strong> (25) &middot; <strong>Co. revenue</strong> (15)</div>
      <div style="padding-left:8px;margin-bottom:2px">Bonus: Comp installs (10) &middot; VJ base (8) &middot; Revenue (6) &middot; Upgrade (4)</div>
      <div style="padding-left:8px;color:var(--text-dim)">Top 25%=A, next=B, next=C, bottom=D</div>

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:1.5px;margin:12px 0 5px">Health Grade (active only)</div>
      <div style="margin-bottom:4px"><em>How strong is the relationship?</em></div>
      <div style="padding-left:8px;margin-bottom:2px"><strong>Rev size</strong> (25) &middot; <strong>Trend</strong> (15) &middot; <strong>YOY %</strong> (15) &middot; <strong>Tech breadth</strong> (15)</div>
      <div style="padding-left:8px;margin-bottom:2px">Contract (10) &middot; Upgrade (8) &middot; Break-fix (7) &middot; Prod lines (10)</div>
      <div style="padding-left:8px;color:var(--text-dim)">A=thriving, D=at-risk</div>

      <div style="color:var(--accent);font-size:9px;text-transform:uppercase;letter-spacing:1.5px;margin:12px 0 5px">Together</div>
      <div style="padding-left:8px;margin-bottom:2px"><strong>Opp A + Health A</strong> &mdash; Protect</div>
      <div style="padding-left:8px;margin-bottom:2px"><strong>Opp A + Health D</strong> &mdash; Needs attention</div>
      <div style="padding-left:8px;margin-bottom:2px"><strong>Opp D + Health A</strong> &mdash; Loyal, low priority</div>
    </div>
    <div class="edit-btns" style="margin-top:12px">
      <div class="edit-btn" onclick="closeRulesPanel()">Close</div>
    </div>
  </div>
</div>

<script>
const RAW_DATA = ${RAW_DATA_JSON};

// ── LocalStorage ──
const LS_KEY = 'vj_territory_edits';
function loadEdits() { try { return JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch(e) { return {}; } }
function saveEdits(edits) { localStorage.setItem(LS_KEY, JSON.stringify(edits)); }
function getEditKey(d) { return (d.company + '|' + (d.address || '')).toLowerCase(); }
function getEffective(d) {
  const edits = loadEdits();
  const key = getEditKey(d);
  const e = edits[key] || {};
  return { ...d, grade: e.grade || d.grade, status: e.status || d.status, star: !!e.star, ring: !!e.ring, note: e.note || '', deleted: !!e.deleted };
}

// ── Map ──
const map = L.map('map', { center:[38.8,-96.5], zoom:6, zoomControl:true });
L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', { maxZoom:19, subdomains:'abcd' }).addTo(map);

let filters = { state:'ALL', grade:'ALL', status:'ALL', vert:'ALL', health:'ALL', upgrade:'ALL', contacts:'ALL', opps:'ALL', rev:'ALL', starred:'ALL', commented:'ALL', search:'' };
let markers = [];
const plainGroup = L.layerGroup().addTo(map);

function formatSales(n) {
  if (!n) return '$0';
  if (n >= 1000000) return '$' + (n/1000000).toFixed(1) + 'M';
  if (n >= 1000) return '$' + Math.round(n/1000) + 'K';
  return '$' + n;
}

function getMarkerClass(d) {
  const eff = getEffective(d);
  if (eff.status === 'active') return 'status-active';
  if (eff.status === 'lapsed') return 'status-lapsed';
  if (eff.status === 'dormant') return 'status-dormant';
  return 'grade-' + eff.grade;
}

function matchesFilters(d) {
  const eff = getEffective(d);
  if (eff.deleted) return false;
  if (filters.state !== 'ALL' && eff.state !== filters.state) return false;
  if (filters.grade !== 'ALL' && eff.grade !== filters.grade) return false;
  if (filters.status !== 'ALL' && eff.status !== filters.status) return false;
  if (filters.vert !== 'ALL') {
    const v = eff.vertical.toLowerCase(), fv = filters.vert.toLowerCase();
    if (fv==='food' && !v.match(/food|meat|poultry|baked|cereal|fruit|vegetable|frozen|snack|candy|confect|flour|grain|sausage/)) return false;
    if (fv==='beverage' && !v.match(/beverage|brew|bottl|distill|wine|beer|coffee|coca|soda/)) return false;
    if (fv==='petfood' && !v.match(/pet|animal feed/)) return false;
    if (fv==='pharma' && !v.match(/pharma|medical|biotech|animal health|cosmetic|personal care/)) return false;
    if (fv==='packaging' && !v.match(/packag|container|carton|label/)) return false;
    if (fv==='aero' && !v.match(/aero|aircraft|aviation|defense|ammunition/)) return false;
    if (fv==='auto' && !v.match(/auto|transport|vehicle|tire|brake/)) return false;
    if (fv==='chemical' && !v.match(/chemical|lubricant|paint|adhesive|coating/)) return false;
    if (fv==='plastics' && !v.match(/plastic|rubber|polymer|extrusion|wire.*cable|cable/)) return false;
    if (fv==='building' && !v.match(/building|construct|hvac|lighting|furniture|hose/)) return false;
    if (fv==='industrial' && !v.match(/industrial|equipment|metal|steel|foundry|manufact/)) return false;
    if (fv==='electronics' && !v.match(/electr|electro|tech/)) return false;
    if (fv==='consumer' && !v.match(/consumer|cpg/)) return false;
    if (fv==='graphics' && !v.match(/graphic|print|text|distribut/)) return false;
    if (fv==='other' && !v.match(/other|unknown|postal|whol/) && !v.match(/^\d/)) return false;
  }
  if (filters.health !== 'ALL') {
    if (filters.health === 'none' && eff.healthGrade) return false;
    if (filters.health !== 'none' && eff.healthGrade !== filters.health) return false;
  }
  if (filters.upgrade !== 'ALL') {
    if (filters.upgrade === 'Y' && !eff.upgradeReady) return false;
    if (filters.upgrade === 'N' && eff.upgradeReady) return false;
  }
  if (filters.contacts !== 'ALL') {
    const cnt = eff.contacts ? eff.contacts.length : 0;
    if (filters.contacts === '0' && cnt !== 0) return false;
    if (filters.contacts === '1' && cnt !== 1) return false;
    if (filters.contacts === '2+' && cnt < 2) return false;
  }
  if (filters.opps !== 'ALL') {
    const hasOpps = eff.oppNote && eff.oppNote.length > 0;
    if (filters.opps === 'Y' && !hasOpps) return false;
    if (filters.opps === 'N' && hasOpps) return false;
  }
  if (filters.rev !== 'ALL') {
    const r = eff.rev012 || 0;
    if (filters.rev === '0' && r !== 0) return false;
    if (filters.rev === '1-10k' && (r <= 0 || r >= 10000)) return false;
    if (filters.rev === '10-50k' && (r < 10000 || r >= 50000)) return false;
    if (filters.rev === '50-100k' && (r < 50000 || r >= 100000)) return false;
    if (filters.rev === '100k+' && r < 100000) return false;
  }
  if (filters.starred !== 'ALL') {
    if (filters.starred === 'Y' && !eff.star) return false;
    if (filters.starred === 'N' && eff.star) return false;
  }
  if (filters.commented !== 'ALL') {
    const hasComments = (eff.note && eff.note.trim().length > 0) || (eff.callNotes && eff.callNotes.length > 0);
    if (filters.commented === 'Y' && !hasComments) return false;
    if (filters.commented === 'N' && hasComments) return false;
  }
  if (filters.search) {
    const s = filters.search.toLowerCase();
    if (!eff.company.toLowerCase().includes(s) && !eff.city.toLowerCase().includes(s) && !eff.vertical.toLowerCase().includes(s) && !eff.state.toLowerCase().includes(s) && !eff.address.toLowerCase().includes(s)) return false;
  }
  return true;
}

function buildPopup(d, idx) {
  const eff = getEffective(d);
  const hasContacts = eff.contacts && eff.contacts.length > 0;
  const statusLabel = eff.status === 'active' ? 'Active Customer' : eff.status === 'lapsed' ? 'Lapsed Customer' : eff.status === 'dormant' ? 'Dormant Customer' : 'Prospect';
  const statusClass = eff.status === 'active' ? 'badge-active' : eff.status === 'lapsed' ? 'badge-lapsed' : eff.status === 'dormant' ? 'badge-dormant' : '';
  const gradeClass = 'badge-' + eff.grade.toLowerCase();
  const trendIcon = eff.trend === 'Rising' ? '<span style="color:var(--a-color)">&uarr; Growing</span>' : eff.trend === 'Declining' ? '<span style="color:#FF6B6B">&darr; Declining</span>' : '&mdash;';

  return \`<div class="popup-inner">
    <div class="popup-company">\${eff.star ? '&#x2B50; ' : ''}\${eff.company}</div>
    <div class="popup-row">&#x1F4CD; \${eff.address ? eff.address + ', ' : ''}\${eff.city}, \${eff.state}</div>
    <div class="popup-badges">
      <span class="badge \${statusClass || 'badge-vert'}">\${statusLabel}</span>
      <span class="badge \${gradeClass}">OPP: \${eff.grade}</span>
      \${eff.healthGrade ? \`<span class="badge badge-\${eff.healthGrade === 'A' ? 'a' : eff.healthGrade === 'B' ? 'b' : eff.healthGrade === 'C' ? 'c' : 'd'}">HEALTH: \${eff.healthGrade}</span>\` : ''}
      <span class="badge badge-vert">\${eff.vertical}</span>
    </div>
    <div class="popup-details">
      \${eff.rev012 ? \`<div class="popup-detail-row"><span>Rev 0-12mo</span><span class="val green">\${formatSales(eff.rev012)}</span></div>\` : ''}
      \${eff.rev1224 ? \`<div class="popup-detail-row"><span>Rev 12-24mo</span><span class="val">\${formatSales(eff.rev1224)}</span></div>\` : ''}
      \${eff.revYOYDiff ? \`<div class="popup-detail-row"><span>YOY Change</span><span class="val \${eff.revYOYDiff > 0 ? 'green' : 'red'}">\${eff.revYOYDiff > 0 ? '+' : ''}\${formatSales(eff.revYOYDiff)}</span></div>\` : ''}
      <div class="popup-detail-row"><span>YOY</span><span class="val">\${trendIcon}</span></div>
      <div class="popup-detail-row"><span>Upgrade</span><span class="val">\${eff.upgradeReady ? 'Yes' : 'No'}</span></div>
      <div class="popup-detail-row"><span>VJ Systems</span><span class="val">\${eff.prodLines}</span></div>
      \${eff.apRev012 ? \`<div class="popup-detail-row"><span>Equip Purchased 0-12mo</span><span class="val green">\${formatSales(eff.apRev012)}</span></div>\` : ''}
      \${eff.apRev1224 ? \`<div class="popup-detail-row"><span>Equip Purchased 12-24mo</span><span class="val">\${formatSales(eff.apRev1224)}</span></div>\` : ''}
      \${eff.vjEquip && Object.keys(eff.vjEquip).length > 0 ? Object.entries(eff.vjEquip).map(([tech, count]) => \`<div class="popup-detail-row" style="padding-left:10px"><span>\${tech}</span><span class="val">\${count}</span></div>\`).join('') : ''}
      <div class="popup-detail-row"><span>Comp CIJ</span><span class="val">\${eff.compCIJ}</span></div>
      <div class="popup-detail-row"><span>Comp Laser</span><span class="val">\${eff.compLaser}</span></div>
      \${eff.oppNote ? \`<div class="popup-detail-row"><span>Opps</span><span class="val">\${eff.oppNote}</span></div>\` : ''}
    </div>
    \${eff.note ? \`<div class="popup-note">\${eff.note}</div>\` : ''}
    <div class="popup-actions">
      \${hasContacts ? \`<div class="popup-btn primary" onclick="showContacts(\${idx})">&#x1F465; Contacts (\${eff.contacts.length})</div>\` : ''}
      <div class="popup-btn" onclick="openEditModal(\${idx})">&#x270F;&#xFE0F; Edit</div>
    </div>
  </div>\`;
}

function render() {
  markers = [];
  plainGroup.clearLayers();

  const edits = loadEdits();
  const customAccounts = [];
  Object.keys(edits).forEach(key => { const e = edits[key]; if (e.isCustom && !e.deleted) customAccounts.push(e.data); });
  const allData = [...RAW_DATA, ...customAccounts];
  const filtered = allData.map((d,i)=>({...d,_idx:i})).filter(matchesFilters);

  let activeCount=0, lapsedCount=0, dormantCount=0, prospectCount=0;
  filtered.forEach(d => {
    const eff = getEffective(d);
    if (eff.status==='active') activeCount++;
    else if (eff.status==='lapsed') lapsedCount++;
    else if (eff.status==='dormant') dormantCount++;
    else prospectCount++;
  });

  document.getElementById('s-total').textContent = filtered.length;
  document.getElementById('s-active').textContent = activeCount;
  document.getElementById('s-lapsed').textContent = lapsedCount;
  document.getElementById('s-dormant').textContent = dormantCount;
  document.getElementById('s-prospect').textContent = prospectCount;

  // Render individual markers (no clustering)
  filtered.forEach(d => {
    const eff = getEffective(d);
    const cls = getMarkerClass(d);
    const starHTML = eff.star ? '<span class="star-overlay">&#x2B50;</span>' : '';
    const ringHTML = eff.ring ? '<span class="ring-overlay"></span>' : '';
    const icon = L.divIcon({
      className:'',
      html:\`<div class="map-marker \${cls}">\${starHTML}\${ringHTML}</div>\`,
      iconSize:[12,12], iconAnchor:[6,6]
    });
    const m = L.marker([d.lat, d.lng], { icon, zIndexOffset: d.status==='active'?200:d.status==='lapsed'?150:d.status==='dormant'?100:d.grade==='A'?50:0 })
      .bindPopup(buildPopup(d, d._idx), { maxWidth:300 });
    m.on('click', () => { document.getElementById('sidebar').classList.remove('open'); document.getElementById('sidebar-overlay').classList.remove('open'); });
    plainGroup.addLayer(m);
    markers.push({ m, idx: d._idx });
  });
}

// ── Sidebar ──
window.toggleSidebar = function() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebar-overlay').classList.toggle('open');
};

function setupPills(groupId, filterKey) {
  document.getElementById(groupId).querySelectorAll('.pill').forEach(p => {
    p.addEventListener('click', () => {
      document.getElementById(groupId).querySelectorAll('.pill').forEach(x => x.classList.remove('active'));
      p.classList.add('active');
      filters[filterKey] = p.dataset.val;
      render();
    });
  });
}
setupPills('f-state','state');
setupPills('f-grade','grade');
setupPills('f-status','status');
setupPills('f-vert','vert');
setupPills('f-health','health');
setupPills('f-upgrade','upgrade');
setupPills('f-contacts','contacts');
setupPills('f-opps','opps');
setupPills('f-rev','rev');
setupPills('f-starred','starred');
setupPills('f-commented','commented');

document.getElementById('search').addEventListener('input', e => { filters.search = e.target.value.trim(); render(); });

window.resetFilters = function() {
  filters = { state:'ALL', grade:'ALL', status:'ALL', vert:'ALL', health:'ALL', upgrade:'ALL', contacts:'ALL', opps:'ALL', rev:'ALL', starred:'ALL', commented:'ALL', search:'' };
  document.getElementById('search').value = '';
  ['f-state','f-grade','f-status','f-vert','f-health','f-upgrade','f-contacts','f-opps','f-rev','f-starred','f-commented'].forEach(id => {
    document.getElementById(id).querySelectorAll('.pill').forEach(p => p.classList.toggle('active', p.dataset.val==='ALL'));
  });
  render();
};

// ── Contacts ──
window.showContacts = function(idx) {
  const edits = loadEdits();
  const customAccounts = [];
  Object.keys(edits).forEach(key => { const e = edits[key]; if (e.isCustom && !e.deleted) customAccounts.push(e.data); });
  const allData = [...RAW_DATA, ...customAccounts];
  const d = allData[idx];
  const eff = getEffective(d);
  document.getElementById('cp-name').textContent = eff.company;
  const list = document.getElementById('cp-list');
  if (!eff.contacts || eff.contacts.length === 0) {
    list.innerHTML = '<div class="cp-empty">No contacts on file</div>';
  } else {
    list.innerHTML = eff.contacts.map(c => {
      const hasPhone = c.phone && c.phone.trim() !== '';
      const hasMobile = c.mobile && c.mobile.trim() !== '';
      const hasEmail = c.email && c.email.trim() !== '';
      const clean = n => n.replace(/[^0-9+]/g, '');
      const isEquip = c.title === 'Equipment Contact';
      const fitClass = isEquip ? 'equip' : (c.fit==='A'||c.fit==='B') ? 'fit-'+c.fit : 'fit-other';
      return \`<div class="cp-contact \${fitClass}">
        <div class="cp-name">\${c.name}\${c.dm?'<span class="cp-dm">&#x2605; DM</span>':''}\${isEquip?'<span class="cp-equip">EQUIP</span>':''}</div>
        <div class="cp-title">\${c.title||'Title unknown'}</div>
        \${hasPhone?\`<div class="cp-phone-row">&#x1F4DE; <a href="tel:+1\${clean(c.phone)}">\${c.phone}</a></div>\`:''}
        \${hasMobile?\`<div class="cp-phone-row">&#x1F4F1; <a href="tel:+1\${clean(c.mobile)}">\${c.mobile}</a></div>\`:''}
        \${hasEmail?\`<div class="cp-email-row">&#x2709; <a href="mailto:\${c.email}">\${c.email}</a></div>\`:''}
        \${!hasPhone&&!hasMobile&&!hasEmail?'<div class="cp-nophone">No contact info on file</div>':''}
      </div>\`;
    }).join('');
  }
  map.closePopup();
  document.getElementById('contacts-panel').classList.add('open');
};

window.closeContacts = function() { document.getElementById('contacts-panel').classList.remove('open'); };

map.on('click', () => { closeContacts(); });

// ── Edit ──
let editingIdx = null;

window.openEditModal = function(idx) {
  editingIdx = idx;
  const edits = loadEdits();
  const customAccounts = [];
  Object.keys(edits).forEach(key => { const e = edits[key]; if (e.isCustom && !e.deleted) customAccounts.push(e.data); });
  const allData = [...RAW_DATA, ...customAccounts];
  const d = allData[idx];
  const eff = getEffective(d);
  document.getElementById('edit-title').textContent = 'Edit: ' + eff.company;
  document.getElementById('edit-star').value = eff.star ? '1' : '0';
  document.getElementById('edit-ring').value = eff.ring ? '1' : '0';
  const key = getEditKey(d);
  const e = edits[key] || {};
  document.getElementById('edit-grade').value = e.grade || '';
  document.getElementById('edit-status').value = e.status || '';
  document.getElementById('edit-note').value = eff.note || '';
  document.getElementById('edit-delete').onclick = function() {
    if (confirm('Delete ' + eff.company + '?')) {
      const edits = loadEdits();
      edits[key] = { ...(edits[key]||{}), deleted: true };
      saveEdits(edits);
      closeEditModal();
      map.closePopup();
      render();
    }
  };
  document.getElementById('edit-modal').classList.add('open');
  map.closePopup();
};

window.saveEdit = function() {
  const edits = loadEdits();
  const customAccounts = [];
  Object.keys(edits).forEach(key => { const e = edits[key]; if (e.isCustom && !e.deleted) customAccounts.push(e.data); });
  const allData = [...RAW_DATA, ...customAccounts];
  const d = allData[editingIdx];
  const key = getEditKey(d);
  const existing = edits[key] || {};
  edits[key] = {
    ...existing,
    star: document.getElementById('edit-star').value === '1',
    ring: document.getElementById('edit-ring').value === '1',
    grade: document.getElementById('edit-grade').value || null,
    status: document.getElementById('edit-status').value || null,
    note: document.getElementById('edit-note').value.trim()
  };
  saveEdits(edits);
  closeEditModal();
  render();
};

window.closeEditModal = function() { document.getElementById('edit-modal').classList.remove('open'); editingIdx = null; };

// ── Add Account ──
let addingPin = false;

window.openAddAccount = function() {
  toggleSidebar();
  document.getElementById('add-name').value = '';
  document.getElementById('add-address').value = '';
  document.getElementById('add-city').value = '';
  document.getElementById('add-state').value = 'KS';
  document.getElementById('add-grade').value = 'C';
  document.getElementById('add-note').value = '';
  document.getElementById('add-lat').value = '';
  document.getElementById('add-lng').value = '';
  document.getElementById('add-modal').classList.add('open');
  addingPin = true;
};

map.on('click', function(e) {
  if (addingPin) {
    document.getElementById('add-lat').value = e.latlng.lat.toFixed(4);
    document.getElementById('add-lng').value = e.latlng.lng.toFixed(4);
  }
});

window.saveNewAccount = function() {
  const name = document.getElementById('add-name').value.trim();
  const lat = parseFloat(document.getElementById('add-lat').value);
  const lng = parseFloat(document.getElementById('add-lng').value);
  if (!name) { alert('Enter a company name'); return; }
  if (!lat || !lng || isNaN(lat) || isNaN(lng)) { alert('Set coordinates first'); return; }
  const newAccount = {
    company: name, address: document.getElementById('add-address').value.trim(),
    city: document.getElementById('add-city').value.trim(), state: document.getElementById('add-state').value,
    vertical: 'Other / Unknown', lat, lng, status: 'prospect',
    grade: document.getElementById('add-grade').value, score: 0,
    rev012: 0, rev1224: 0, trend: 'unknown', upgradeReady: false,
    compCIJ: 0, compLaser: 0, prodLines: 0, oppNote: '', contacts: []
  };
  const edits = loadEdits();
  const key = getEditKey(newAccount);
  edits[key] = { isCustom: true, data: newAccount, note: document.getElementById('add-note').value.trim(), star: false, ring: false };
  saveEdits(edits);
  closeAddModal();
  render();
};

window.closeAddModal = function() { document.getElementById('add-modal').classList.remove('open'); addingPin = false; };

window.openRulesPanel = function() { document.getElementById('rules-panel').classList.add('open'); };
window.closeRulesPanel = function() { document.getElementById('rules-panel').classList.remove('open'); };

window.exportEdits = function() {
  const edits = loadEdits();
  const blob = new Blob([JSON.stringify(edits, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url;
  a.download = 'vj-territory-edits-' + new Date().toISOString().slice(0,10) + '.json';
  a.click(); URL.revokeObjectURL(url);
};

render();
setTimeout(() => map.invalidateSize(), 100);
<\/script>
</body>
</html>`;

// ── Write files ──
fs.writeFileSync('index.html', desktopHTML);
fs.writeFileSync('vj-territory-mobile.html', mobileHTML);

console.log('\\nDone! Wrote index.html and vj-territory-mobile.html');
console.log(`Total accounts in RAW_DATA: ${accounts.length}`);
