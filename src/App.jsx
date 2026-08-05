import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────
const DEFAULT_CATS = {
  personal: {
    income: ["Salary", "Freelance", "Investment", "Gift", "Other"],
    expense: [
      "Ex-0011: Phone",
      "Ex-0012: Bank/Transfer/Visa Fee",
      "Ex-0014: Investment Fund (Stock)",
      "Ex-0015: School Fee",
      "Ex-0016: Loan Payment",
      "Ex-0017: Insurance",
      "Ex-0018: Lost or Unadjustable",
      "Ex-0021: Main Meals for Family",
      "Ex-0022: Party & Ceremony",
      "Ex-0023: Eat Out/Date",
      "Ex-0024: Snack & Drink",
      "Ex-0031: Petrol and Gas",
      "Ex-0033: Transportation",
      "Ex-0041: Life Event and Trip",
      "Ex-0042: Hobby & Movie",
      "Ex-0043: Education & Books",
      "Ex-0044: Charity",
      "Ex-0045: Gifts/Jewelry",
      "Ex-0051: Skincare and Hair",
      "Ex-0052: Stationary",
      "Ex-0054: Cloth & Bag",
      "Ex-0061: Health Care",
      "Ex-0062: Therapist and Massage",
      "Ex-0065: Health Gear",
      "Ex-0071: Appliances/Items",
      "Ex-0073: TV and Internet",
      "Ex-0082: Car Expenses",
      "Ex-0083: Coffee Shop Expenses",
      "Other",
    ],
  },
  coffeeshop: {
    income: ["Coffee Sales", "Food Sales", "Catering", "Tips", "Other"],
    expense: ["Beans & Supplies", "Rent", "Staff Wages", "Equipment", "Utilities", "Marketing", "Other"],
  },
};

const POCKET_META = {
  personal:   { id: "personal",   label: "Personal",    emoji: "🏠", accent: "#1a1a1a", light: "#f7f5f2" },
  coffeeshop: { id: "coffeeshop", label: "Coffee Shop",  emoji: "☕", accent: "#6B3E26", light: "#fdf8f4" },
};

const DEFAULT_BUDGETS = {
  personal: [
    { category: "Food", limit: 500, color: "#E8C547" },
    { category: "Transport", limit: 200, color: "#7EC8A4" },
    { category: "Housing", limit: 1500, color: "#F4A07A" },
    { category: "Shopping", limit: 300, color: "#A8C5E8" },
    { category: "Entertainment", limit: 150, color: "#C5A8E8" },
  ],
  coffeeshop: [
    { category: "Beans & Supplies", limit: 800,  color: "#C8916B" },
    { category: "Staff Wages",       limit: 3000, color: "#7EC8A4" },
    { category: "Rent",              limit: 2000, color: "#F4A07A" },
    { category: "Utilities",         limit: 400,  color: "#A8C5E8" },
    { category: "Marketing",         limit: 200,  color: "#C5A8E8" },
  ],
};

const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const COLORS  = ["#E8C547","#7EC8A4","#F4A07A","#A8C5E8","#C5A8E8","#C8916B","#E8A4A4","#A4C8E8","#B4E8A4","#E8B4A4"];

const fmt   = n => new Intl.NumberFormat("en-US",{ style:"currency", currency:"USD", maximumFractionDigits:0 }).format(n);
const today = () => new Date().toISOString().split("T")[0];

// ─── Seed Data (Heang's real data restored from backup) ──────────────────────
const SEED = {
  personal: [
    { id:1,   type:"income",  category:"Salary",           amount:962.62, note:"May salary",                    date:"2026-05-30" },
    { id:2,   type:"income",  category:"Salary",           amount:962.54, note:"",                              date:"2026-06-30" },
    { id:3,   type:"income",  category:"Other",            amount:15,     note:"Sister own",                    date:"2026-06-03" },
    { id:4,   type:"expense", category:"Other",            amount:4,      note:"Netflix subscription",          date:"2026-06-01" },
    { id:5,   type:"expense", category:"Food",             amount:6.25,   note:"",                              date:"2026-06-03" },
    { id:6,   type:"expense", category:"Food",             amount:2,      note:"",                              date:"2026-06-04" },
    { id:7,   type:"expense", category:"Trip",             amount:13.58,  note:"Trip to pp",                    date:"2026-06-04" },
    { id:8,   type:"expense", category:"Event/Charity",    amount:10,     note:"Contribute with bong Vanthat",  date:"2026-06-04" },
    { id:9,   type:"expense", category:"Trip",             amount:96,     note:"Trip to Chiphat",               date:"2026-06-07" },
    { id:10,  type:"expense", category:"Event/Charity",    amount:17,     note:"Wedding Reaksmey",              date:"2026-06-07" },
    { id:11,  type:"expense", category:"Health",           amount:19,     note:"Pp trip for doctor appointment",date:"2026-06-09" },
    { id:12,  type:"expense", category:"Food",             amount:3.25,   note:"Drinks",                        date:"2026-06-09" },
    { id:13,  type:"expense", category:"Event/Charity",    amount:2.5,    note:"Donations",                     date:"2026-06-10" },
    { id:14,  type:"expense", category:"Snack and drink",  amount:2,      note:"Drinks",                        date:"2026-06-11" },
    { id:15,  type:"expense", category:"Health",           amount:1,      note:"Exercise",                      date:"2026-06-11" },
    { id:16,  type:"expense", category:"Item",             amount:4,      note:"Softer",                        date:"2026-06-13" },
    { id:17,  type:"expense", category:"Shopping",         amount:3.5,    note:"T-shirt",                       date:"2026-06-14" },
    { id:18,  type:"expense", category:"Health",           amount:1.5,    note:"Exercise",                      date:"2026-06-15" },
    { id:19,  type:"expense", category:"Event/Charity",    amount:10,     note:"Help Ms. oun",                  date:"2026-06-15" },
    { id:20,  type:"expense", category:"Transport",        amount:1.2,    note:"Gasoline",                      date:"2026-06-15" },
    { id:21,  type:"expense", category:"Item",             amount:4.5,    note:"Water bottle",                  date:"2026-06-17" },
    { id:22,  type:"expense", category:"Shopping",         amount:10,     note:"Clothes",                       date:"2026-06-17" },
    { id:23,  type:"expense", category:"Item",             amount:10.16,  note:"Necklace",                      date:"2026-06-18" },
    { id:24,  type:"expense", category:"Trip",             amount:87.34,  note:"Trip to kampot",                date:"2026-06-22" },
    { id:25,  type:"expense", category:"Trip",             amount:22.54,  note:"Other expenses during trip to Kampot", date:"2026-06-22" },
    { id:26,  type:"expense", category:"Skin care/Hair",   amount:21.5,   note:"Skin care",                     date:"2026-06-22" },
    { id:27,  type:"expense", category:"Health",           amount:1,      note:"Exercise",                      date:"2026-06-22" },
    { id:28,  type:"income",  category:"Salary",           amount:100,    note:"Uncle Phorn gave",              date:"2026-06-23" },
    { id:29,  type:"expense", category:"Food",             amount:6.87,   note:"Vegetable",                     date:"2026-06-23" },
    { id:30,  type:"expense", category:"Transport",        amount:2.2,    note:"Gasoline",                      date:"2026-06-23" },
    { id:31,  type:"expense", category:"Snack and drink",  amount:1.25,   note:"",                              date:"2026-06-23" },
    { id:32,  type:"expense", category:"Phone",            amount:4,      note:"Phone card",                    date:"2026-06-23" },
    { id:33,  type:"expense", category:"Jewelry/gold",     amount:992,    note:"2G gold",                       date:"2026-06-24" },
    { id:34,  type:"expense", category:"Health",           amount:1,      note:"Exercise",                      date:"2026-06-25" },
    { id:35,  type:"expense", category:"Skin care/Hair",   amount:5,      note:"Lipstick",                      date:"2026-06-25" },
    { id:36,  type:"expense", category:"Food",             amount:4,      note:"",                              date:"2026-06-27" },
    { id:37,  type:"expense", category:"Food",             amount:4.5,    note:"For family",                    date:"2026-06-27" },
    { id:38,  type:"expense", category:"Food",             amount:6,      note:"",                              date:"2026-06-29" },
    { id:39,  type:"income",  category:"Other",            amount:32,     note:"Owe from Navy",                 date:"2026-06-30" },
    { id:40,  type:"expense", category:"Snack and drink",  amount:5.68,   note:"Snack farewell",                date:"2026-06-30" },
    { id:41,  type:"expense", category:"Item",             amount:2.5,    note:"Washing car",                   date:"2026-06-30" },
    { id:42,  type:"expense", category:"Skin care/Hair",   amount:12,     note:"Shampoo",                       date:"2026-06-30" },
    { id:43,  type:"expense", category:"Shopping",         amount:1.3,    note:"",                              date:"2026-06-05" },
    { id:44,  type:"expense", category:"Food",             amount:2.5,    note:"Exercise and avocado",          date:"2026-07-03" },
    { id:45,  type:"expense", category:"Item",             amount:1,      note:"",                              date:"2026-07-03" },
    { id:46,  type:"expense", category:"Food",             amount:2.25,   note:"Breakfast",                     date:"2026-07-04" },
    { id:47,  type:"expense", category:"Food",             amount:2,      note:"",                              date:"2026-07-07" },
    { id:48,  type:"expense", category:"Health",           amount:36.65,  note:"Trip to pp",                    date:"2026-07-07" },
    { id:49,  type:"expense", category:"Item",             amount:1.5,    note:"Exercise and detergent",        date:"2026-07-09" },
    { id:50,  type:"expense", category:"Food",             amount:6.25,   note:"Eat out",                       date:"2026-07-11" },
    { id:51,  type:"expense", category:"Health",           amount:0.75,   note:"Exercise",                      date:"2026-07-11" },
    { id:52,  type:"expense", category:"Health",           amount:677.7,  note:"Annual insurance fee",          date:"2026-07-13" },
    { id:53,  type:"expense", category:"Transport",        amount:17,     note:"Gasoline car",                  date:"2026-07-13" },
    { id:54,  type:"expense", category:"Item",             amount:3,      note:"Socks",                         date:"2026-07-13" },
    { id:55,  type:"expense", category:"Food",             amount:2,      note:"",                              date:"2026-07-15" },
    { id:56,  type:"expense", category:"Food",             amount:6,      note:"Eat out",                       date:"2026-07-16" },
    { id:57,  type:"expense", category:"Transport",        amount:2,      note:"Gasoline",                      date:"2026-07-16" },
    { id:58,  type:"expense", category:"Health",           amount:23,     note:"Turmeric for mom",              date:"2026-07-16" },
    { id:59,  type:"expense", category:"Transport",        amount:2.12,   note:"Gasoline",                      date:"2026-07-23" },
    { id:60,  type:"expense", category:"Food",             amount:3,      note:"",                              date:"2026-07-22" },
    { id:61,  type:"expense", category:"Entertainment",    amount:4,      note:"Netflix",                       date:"2026-07-22" },
    { id:62,  type:"expense", category:"Item",             amount:16.52,  note:"Taopao",                        date:"2026-07-22" },
    { id:63,  type:"expense", category:"Food",             amount:6.25,   note:"",                              date:"2026-07-19" },
    { id:64,  type:"expense", category:"Jewelry/gold",     amount:571,    note:"Bought necklace and earring",   date:"2026-07-19" },
    { id:65,  type:"income",  category:"Salary",           amount:960,    note:"Advanced salary",               date:"2026-07-19" },
    { id:66,  type:"expense", category:"Food",             amount:4.12,   note:"Snack",                         date:"2026-07-26" },
    { id:67,  type:"expense", category:"Food",             amount:2,      note:"",                              date:"2026-07-24" },
    { id:68,  type:"expense", category:"Item",             amount:76.5,   note:"Laser for hair removal",        date:"2026-07-25" },
    { id:69,  type:"expense", category:"Food",             amount:2,      note:"",                              date:"2026-07-25" },
    { id:70,  type:"expense", category:"Health",           amount:1.5,    note:"Exercise",                      date:"2026-07-27" },
    { id:71,  type:"expense", category:"Food",             amount:3.5,    note:"",                              date:"2026-07-27" },
    { id:72,  type:"expense", category:"Event/Charity",    amount:7,      note:"",                              date:"2026-07-27" },
    { id:73,  type:"expense", category:"Transport",        amount:3.26,   note:"Gasoline + washing my bike",    date:"2026-07-30" },
    { id:74,  type:"expense", category:"Food",             amount:5.87,   note:"",                              date:"2026-07-29" },
    { id:75,  type:"income",  category:"Salary",           amount:164,    note:"For August budget",             date:"2026-07-31" },
    { id:76,  type:"expense", category:"Event/Charity",    amount:30,     note:"Donation to ODC",               date:"2026-07-31" },
    { id:77,  type:"expense", category:"Food",             amount:16,     note:"Durian",                        date:"2026-08-02" },
    { id:78,  type:"expense", category:"Skin care/Hair",   amount:5,      note:"",                              date:"2026-08-02" },
  ],
  coffeeshop: [
    { id:101, type:"income",  category:"Coffee Sales",     amount:2412.30, note:"ABA bank",                    date:"2026-05-31" },
    { id:102, type:"income",  category:"Coffee Sales",     amount:693.27,  note:"ACLEDA BANK",                 date:"2026-05-31" },
    { id:103, type:"income",  category:"Other",            amount:3805.60, note:"",                            date:"2026-05-31" },
    { id:104, type:"income",  category:"Tips",             amount:843.67,  note:"From mama",                   date:"2026-05-31" },
    { id:105, type:"expense", category:"Marketing",        amount:100,     note:"Sent to mom account",         date:"2026-05-31" },
    { id:106, type:"expense", category:"Equipment",        amount:3106.54, note:"May total expense",           date:"2026-05-31" },
    { id:107, type:"income",  category:"Coffee Sales",     amount:2445.29, note:"",                            date:"2026-06-30" },
    { id:108, type:"expense", category:"Other",            amount:61.94,   note:"Paid NSSF",                   date:"2026-06-01" },
    { id:109, type:"expense", category:"Equipment",        amount:80,      note:"Bag",                         date:"2026-06-08" },
    { id:110, type:"expense", category:"Marketing",        amount:2.8,     note:"Boost TikTok",                date:"2026-06-09" },
    { id:111, type:"expense", category:"Beans & Supplies", amount:332,     note:"Funan coffee",                date:"2026-06-10" },
    { id:112, type:"expense", category:"Equipment",        amount:673.5,   note:"E sour",                      date:"2026-06-12" },
    { id:113, type:"expense", category:"Equipment",        amount:652.5,   note:"E sour",                      date:"2026-06-12" },
    { id:114, type:"expense", category:"Beans & Supplies", amount:57,      note:"Cream powder 10kg",           date:"2026-06-13" },
    { id:115, type:"expense", category:"Utilities",        amount:66.25,   note:"Utility at home",             date:"2026-06-16" },
    { id:116, type:"expense", category:"Beans & Supplies", amount:332,     note:"Funan coffee",                date:"2026-06-24" },
    { id:117, type:"expense", category:"Equipment",        amount:14,      note:"Syrub and cream powder",      date:"2026-06-27" },
    { id:118, type:"income",  category:"Coffee Sales",     amount:2845.29, note:"Sales in June",               date:"2026-07-04" },
    { id:119, type:"expense", category:"Equipment",        amount:666,     note:"Milk",                        date:"2026-07-01" },
    { id:120, type:"expense", category:"Equipment",        amount:646.84,  note:"E sour",                      date:"2026-07-04" },
    { id:121, type:"expense", category:"Equipment",        amount:570,     note:"E sour",                      date:"2026-07-04" },
    { id:122, type:"expense", category:"Beans & Supplies", amount:332,     note:"Funan coffee",                date:"2026-07-09" },
    { id:123, type:"expense", category:"Utilities",        amount:66,      note:"Electricity at home",         date:"2026-07-16" },
    { id:124, type:"expense", category:"Beans & Supplies", amount:332,     note:"Funan coffee",                date:"2026-07-21" },
    { id:125, type:"expense", category:"Equipment",        amount:10,      note:"Paper cup",                   date:"2026-07-21" },
    { id:126, type:"expense", category:"Equipment",        amount:106,     note:"Plastic bags",                date:"2026-07-23" },
    { id:127, type:"expense", category:"Beans & Supplies", amount:276,     note:"Fresh milk 20 boxes",         date:"2026-07-31" },
  ],
};

// ─── Excel Export (with period filter) ───────────────────────────────────────
function exportXLSX(transactions, pocketLabel, periodLabel, filterFn) {
  const wb = XLSX.utils.book_new();
  const filtered = transactions.filter(filterFn).sort((a,b)=>new Date(a.date)-new Date(b.date));

  // ── Sheet 1: Summary ──
  const cashIn  = filtered.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
  const cashOut = filtered.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
  const summaryRows = [
    ["Report Period", periodLabel],
    ["Pocket", pocketLabel],
    ["Generated", new Date().toLocaleDateString()],
    [""],
    ["", "Cash In ($)", "Cash Out ($)", "Net ($)"],
    ["TOTAL", cashIn, cashOut, cashIn - cashOut],
  ];
  const ws0 = XLSX.utils.aoa_to_sheet(summaryRows);
  ws0["!cols"] = [{ wch:20 },{ wch:14 },{ wch:14 },{ wch:14 }];
  XLSX.utils.book_append_sheet(wb, ws0, "Summary");

  // ── Sheet 2: All Transactions (Cash In | Cash Out columns) ──
  const txRows = [["Date", "Category", "Note", "Cash In ($)", "Cash Out ($)"]];
  filtered.forEach(t => {
    txRows.push([
      t.date, t.category, t.note,
      t.type === "income"  ? t.amount : "",
      t.type === "expense" ? t.amount : "",
    ]);
  });
  txRows.push(["", "TOTAL", "", cashIn, cashOut]);
  const ws1 = XLSX.utils.aoa_to_sheet(txRows);
  ws1["!cols"] = [{ wch:12 },{ wch:30 },{ wch:28 },{ wch:14 },{ wch:14 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Transactions");

  // ── Sheet 3: Cash Out by Category ──
  const catSpend = {};
  filtered.filter(t=>t.type==="expense").forEach(t=>{ catSpend[t.category]=(catSpend[t.category]||0)+t.amount; });
  const catRows = [["Category", "Cash Out ($)", "% of Total"]];
  Object.entries(catSpend).sort((a,b)=>b[1]-a[1])
    .forEach(([cat,amt]) => catRows.push([cat, amt, `${((amt/(cashOut||1))*100).toFixed(1)}%`]));
  const ws2 = XLSX.utils.aoa_to_sheet(catRows);
  ws2["!cols"] = [{ wch:32 },{ wch:14 },{ wch:14 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Cash Out by Category");

  // ── Sheet 4: Cash In by Category ──
  const incCat = {};
  filtered.filter(t=>t.type==="income").forEach(t=>{ incCat[t.category]=(incCat[t.category]||0)+t.amount; });
  const incRows = [["Category", "Cash In ($)", "% of Total"]];
  Object.entries(incCat).sort((a,b)=>b[1]-a[1])
    .forEach(([cat,amt]) => incRows.push([cat, amt, `${((amt/(cashIn||1))*100).toFixed(1)}%`]));
  const ws3 = XLSX.utils.aoa_to_sheet(incRows);
  ws3["!cols"] = [{ wch:24 },{ wch:14 },{ wch:14 }];
  XLSX.utils.book_append_sheet(wb, ws3, "Cash In by Category");

  XLSX.writeFile(wb, `${pocketLabel}-${periodLabel.replace(/\s/g,"-")}-report.xlsx`);
}

function exportYearlyXLSX(transactions, pocketLabel) {
  const wb = XLSX.utils.book_new();
  const years = [...new Set(transactions.map(t => new Date(t.date).getFullYear()))].sort((a,b)=>a-b);

  const yearRows = [["Year","Cash In ($)","Cash Out ($)","Net ($)"]];
  years.forEach(y => {
    const yTxs = transactions.filter(t => new Date(t.date).getFullYear()===y);
    const inc  = yTxs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
    const exp  = yTxs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
    yearRows.push([y, inc, exp, inc-exp]);
  });
  const ws1 = XLSX.utils.aoa_to_sheet(yearRows);
  ws1["!cols"] = [{ wch:8 },{ wch:14 },{ wch:14 },{ wch:14 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Year over Year");

  if (years.length > 1) {
    const growthRows = [["Period","Revenue Growth","Expense Growth"]];
    for (let i = 1; i < years.length; i++) {
      const py = years[i-1], cy = years[i];
      const pt = transactions.filter(t=>new Date(t.date).getFullYear()===py);
      const ct = transactions.filter(t=>new Date(t.date).getFullYear()===cy);
      const pInc = pt.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
      const cInc = ct.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
      const pExp = pt.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
      const cExp = ct.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
      const ig = pInc ? `${((cInc-pInc)/pInc*100).toFixed(1)}%` : "—";
      const eg = pExp ? `${((cExp-pExp)/pExp*100).toFixed(1)}%` : "—";
      growthRows.push([`${py} → ${cy}`, ig, eg]);
    }
    const ws2 = XLSX.utils.aoa_to_sheet(growthRows);
    ws2["!cols"] = [{ wch:14 },{ wch:16 },{ wch:16 }];
    XLSX.utils.book_append_sheet(wb, ws2, "Growth");
  }

  XLSX.writeFile(wb, `${pocketLabel}-yearly-report.xlsx`);
}

// ─── PDF Export ───────────────────────────────────────────────────────────────
function exportPDF(title, htmlContent) {
  const win = window.open("","_blank");
  win.document.write(`<!DOCTYPE html><html><head><title>${title}</title>
  <style>
    body{font-family:'Helvetica Neue',Arial,sans-serif;margin:40px;color:#1a1a1a;font-size:13px}
    h1{font-size:22px;margin-bottom:4px}h2{font-size:14px;margin:24px 0 8px;color:#555;border-bottom:1px solid #eee;padding-bottom:5px}
    table{width:100%;border-collapse:collapse;margin-bottom:16px}
    th{background:#f5f5f5;padding:7px 10px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#777}
    td{padding:7px 10px;border-bottom:1px solid #f0f0f0}
    .g{color:#3a8a5e}.r{color:#b8962a}.n{color:#1a1a1a}.nb{color:#c0392b}
    .cards{display:flex;gap:20px;margin-bottom:24px}
    .card{border:1px solid #eee;border-radius:8px;padding:12px 18px}
    .cl{font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:#999;margin-bottom:4px}
    .cv{font-size:20px;font-weight:700}
    @media print{body{margin:20px}}
  </style></head><body>${htmlContent}<script>window.print();<\/script></body></html>`);
  win.document.close();
}

// ─── localStorage helpers ─────────────────────────────────────────────────────
function loadStorage(key, fallback) {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; }
}
function saveStorage(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch {}
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [pocket,    setPocket]   = useState("personal");
  const [tab,       setTab]      = useState("overview");
  const [txData,    setTxData]   = useState(() => loadStorage("fin_txData",  SEED));
  const [budgets,   setBudgets]  = useState(() => loadStorage("fin_budgets", DEFAULT_BUDGETS));
  const [cats,      setCats]     = useState(() => loadStorage("fin_cats",    DEFAULT_CATS));
  const [userName,  setUserName] = useState(() => loadStorage("fin_userName", null));
  const [nameInput, setNameInput]= useState("");
  const [shopInput, setShopInput]= useState("");
  const [shopName,  setShopName] = useState(() => loadStorage("fin_shopName", "Coffee Shop"));
  const [showAdd,      setShowAdd]     = useState(false);
  const [showCats,     setShowCats]    = useState(false);
  const [showSettings, setShowSettings]= useState(false);
  const [form,      setForm]     = useState({ type:"expense", category:"", amount:"", note:"", date:today() });
  const [mounted,   setMounted]  = useState(false);

  // Persist to localStorage whenever data changes
  useEffect(() => { saveStorage("fin_txData",  txData);    }, [txData]);
  useEffect(() => { saveStorage("fin_budgets", budgets);   }, [budgets]);
  useEffect(() => { saveStorage("fin_cats",    cats);      }, [cats]);
  useEffect(() => { saveStorage("fin_userName", userName); }, [userName]);
  useEffect(() => { saveStorage("fin_shopName", shopName); }, [shopName]);

  useEffect(()=>{ setMounted(true); },[]);
  useEffect(()=>{
    setForm(f=>({ ...f, category: cats[pocket][f.type][0] || "" }));
  },[pocket, cats]);

  // ── First-time setup screen ──
  if (!userName) {
    return (
      <div style={s.shell}>
        <style>{css}</style>
        <div style={{ ...s.app, background:"#f7f5f2", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:"100vh", padding:"40px 32px" }}>
          <div style={{ fontSize:48, marginBottom:16 }}>👋</div>
          <div style={{ fontSize:26, fontWeight:700, color:"#1a1a1a", marginBottom:8, textAlign:"center" }}>Welcome!</div>
          <div style={{ fontSize:14, color:"#aaa", marginBottom:40, textAlign:"center" }}>Let's set up your finance tracker.</div>
          <div style={{ width:"100%", maxWidth:360 }}>
            <div style={{ fontSize:12, fontWeight:700, color:"#bbb", textTransform:"uppercase", letterSpacing:1.2, marginBottom:8 }}>Your Name</div>
            <input
              style={{ ...s.input, fontSize:17, marginBottom:6 }}
              placeholder="e.g. Heang"
              value={nameInput}
              onChange={e => setNameInput(e.target.value)}
            />
            <div style={{ fontSize:12, color:"#ccc", marginBottom:24 }}>This will appear as "<b>{nameInput.trim() || "Your name"}'s Wallet</b>"</div>
            <div style={{ fontSize:12, fontWeight:700, color:"#bbb", textTransform:"uppercase", letterSpacing:1.2, marginBottom:8 }}>Coffee Shop Name</div>
            <input
              style={{ ...s.input, fontSize:17, marginBottom:6 }}
              placeholder="e.g. Heang's Café"
              value={shopInput}
              onChange={e => setShopInput(e.target.value)}
            />
            <div style={{ fontSize:12, color:"#ccc", marginBottom:32 }}>Leave blank to use "Coffee Shop"</div>
            <button
              style={{ ...s.submitBtn, background: nameInput.trim() ? "#1a1a1a" : "#ccc", fontSize:17 }}
              disabled={!nameInput.trim()}
              onClick={() => { if(nameInput.trim()) { setUserName(nameInput.trim()); setShopName(shopInput.trim() || "Coffee Shop"); }}}
            >
              Get Started →
            </button>
          </div>
        </div>
      </div>
    );
  }

  // Dynamic pocket labels using user's name
  const POCKET_DISPLAY = {
    personal:   { ...POCKET_META.personal,   label: `${userName}'s Wallet` },
    coffeeshop: { ...POCKET_META.coffeeshop, label: shopName },
  };

  const P       = POCKET_DISPLAY[pocket];
  const txs     = txData[pocket];
  const pCats   = cats[pocket];
  const pBudgets= budgets[pocket];

  const totalIncome  = txs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
  const totalExpense = txs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
  const balance      = totalIncome - totalExpense;

  const spentByCategory = (cat) =>
    txs.filter(t=>t.type==="expense"&&t.category===cat).reduce((s,t)=>s+t.amount,0);

  const addTransaction = () => {
    if (!form.amount || isNaN(form.amount)) return;
    const newTx = { id: Date.now(), ...form, amount: parseFloat(form.amount) };
    setTxData(prev => ({
      ...prev,
      [pocket]: [...prev[pocket], newTx].sort((a,b) => new Date(b.date) - new Date(a.date))
    }));
    setForm(f=>({ ...f, amount:"", note:"", date:today() }));
    setShowAdd(false);
  };

  const deleteTx = id => setTxData(prev=>({ ...prev, [pocket]:prev[pocket].filter(t=>t.id!==id) }));

  const addCategory = (type, name) => {
    if (!name.trim()) return;
    setCats(prev=>({
      ...prev,
      [pocket]:{ ...prev[pocket], [type]: [...prev[pocket][type].filter(c=>c!=="Other"), name.trim(), "Other"] }
    }));
  };

  const removeCategory = (type, name) => {
    if (["Other"].includes(name)) return;
    setCats(prev=>({ ...prev, [pocket]:{ ...prev[pocket], [type]: prev[pocket][type].filter(c=>c!==name) } }));
  };

  const recent = [...txs].sort((a,b)=>new Date(b.date)-new Date(a.date)).slice(0,6);

  return (
    <div style={s.shell}>
      <style>{css}</style>
      <div style={{ ...s.app, background:P.light }} className={mounted?"mounted":""}>

        {/* Pocket switcher */}
        <div style={s.pocketBar}>
          {Object.values(POCKET_DISPLAY).map(p=>(
            <button key={p.id}
              style={{ ...s.pocketBtn, ...(pocket===p.id?{ ...s.pocketActive, background:p.accent }:{}) }}
              onClick={()=>{ setPocket(p.id); setTab("overview"); }}>
              <span>{p.emoji}</span><span>{p.label}</span>
            </button>
          ))}
        </div>

        {/* Header */}
        <div style={s.header}>
          <div>
            <div style={s.headerLabel}>{P.label}</div>
            <div style={s.headerSub}>All time · {txs.length} transactions</div>
          </div>
          <div style={{ display:"flex", gap:8 }}>
            <button style={{ ...s.addBtn, background:"#f0ede8", fontSize:18 }} onClick={()=>setShowSettings(true)}>⚙️</button>
            <button style={{ ...s.addBtn, background:P.accent }} onClick={()=>setShowAdd(true)}>+</button>
          </div>
        </div>

        {/* Balance card */}
        <div style={{ ...s.balanceCard, borderTop:`4px solid ${P.accent}` }} className="fade-in">
          <div style={s.balanceLabel}>Net Balance</div>
          <div style={{ ...s.balanceAmount, color:balance>=0?P.accent:"#E05A5A" }}>{fmt(balance)}</div>
          <div style={s.balanceRow}>
            <div style={s.balanceStat}>
              <span style={{ ...s.dot, background:"#7EC8A4" }}/><span style={s.statLabel}>Income</span>
              <span style={s.statVal}>{fmt(totalIncome)}</span>
            </div>
            <div style={s.balanceDivider}/>
            <div style={s.balanceStat}>
              <span style={{ ...s.dot, background:"#E8C547" }}/><span style={s.statLabel}>Spent</span>
              <span style={s.statVal}>{fmt(totalExpense)}</span>
            </div>
          </div>
        </div>

        {/* Tabs */}
        <div style={s.tabs}>
          {["overview","transactions","budgets","reports"].map(t=>(
            <button key={t}
              style={{ ...s.tabBtn, ...(tab===t?{ ...s.tabActive, background:P.accent }:{}) }}
              onClick={()=>setTab(t)}>
              {t==="overview"?"Home":t.charAt(0).toUpperCase()+t.slice(1)}
            </button>
          ))}
        </div>

        <div style={s.content}>
          {tab==="overview"     && <Overview recent={recent} budgets={pBudgets} spentByCategory={spentByCategory} fmt={fmt} deleteTransaction={deleteTx} accent={P.accent}/>}
          {tab==="transactions" && <Transactions transactions={txs} fmt={fmt} deleteTransaction={deleteTx} accent={P.accent}/>}
          {tab==="budgets"      && <Budgets budgets={pBudgets} setBudgets={b=>setBudgets(prev=>({ ...prev, [pocket]:b }))} spentByCategory={spentByCategory} fmt={fmt} accent={P.accent} pocketCats={pCats}/>}
          {tab==="reports"      && <Reports transactions={txs} fmt={fmt} accent={P.accent} pocketLabel={P.label}/>}
        </div>
      </div>

      {/* ── Add Transaction Modal ── */}
      {showAdd && (
        <div style={s.overlay} onClick={()=>setShowAdd(false)}>
          <div style={s.modal} onClick={e=>e.stopPropagation()}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:18 }}>
              <div style={s.modalTitle}>Add to {P.label} {P.emoji}</div>
              <button style={s.manageBtn} onClick={()=>{ setShowAdd(false); setShowCats(true); }}>Manage categories</button>
            </div>
            <div style={s.typeToggle}>
              {["expense","income"].map(t=>(
                <button key={t}
                  style={{ ...s.typeBtn, ...(form.type===t?{ background:t==="income"?"#7EC8A4":"#E8C547", color:"#1a1a1a", border:"none" }:{}) }}
                  onClick={()=>setForm(f=>({ ...f, type:t, category:pCats[t][0]||"" }))}>
                  {t==="income"?"↑ Income":"↓ Expense"}
                </button>
              ))}
            </div>
            <input style={s.input} type="number" placeholder="Amount ($)" value={form.amount} onChange={e=>setForm(f=>({ ...f, amount:e.target.value }))}/>
            <select style={s.input} value={form.category} onChange={e=>setForm(f=>({ ...f, category:e.target.value }))}>
              {pCats[form.type].map(c=><option key={c}>{c}</option>)}
            </select>
            <input style={s.input} placeholder="Note (optional)" value={form.note} onChange={e=>setForm(f=>({ ...f, note:e.target.value }))}/>
            <input style={s.input} type="date" value={form.date} onChange={e=>setForm(f=>({ ...f, date:e.target.value }))}/>
            <button style={{ ...s.submitBtn, background:P.accent }} onClick={addTransaction}>Add Transaction</button>
          </div>
        </div>
      )}

      {/* ── Manage Categories Modal ── */}
      {showCats && (
        <div style={s.overlay} onClick={()=>setShowCats(false)}>
          <div style={{ ...s.modal, maxHeight:"85vh", overflowY:"auto" }} onClick={e=>e.stopPropagation()}>
            <div style={s.modalTitle}>Categories — {P.label} {P.emoji}</div>
            <div style={{ fontSize:12, color:"#aaa", marginBottom:20 }}>Add or remove custom categories for this pocket.</div>
            {["income","expense"].map(type=>(
              <CategoryEditor key={type} type={type} categories={pCats[type]}
                onAdd={name=>addCategory(type,name)} onRemove={name=>removeCategory(type,name)} accent={P.accent}/>
            ))}
            <button style={{ ...s.submitBtn, background:P.accent, marginTop:8 }} onClick={()=>setShowCats(false)}>Done</button>
          </div>
        </div>
      )}

      {/* ── Settings Modal ── */}
      {showSettings && (
        <div style={s.overlay} onClick={()=>setShowSettings(false)}>
          <div style={s.modal} onClick={e=>e.stopPropagation()}>
            <div style={s.modalTitle}>Settings ⚙️</div>
            <div style={{ fontSize:12, color:"#aaa", marginBottom:24 }}>Manage your app data and preferences.</div>

            {/* Reset this pocket */}
            <div style={s.settingsSection}>
              <div style={s.settingsLabel}>Reset "{P.label}" Data</div>
              <div style={{ fontSize:12, color:"#aaa", marginBottom:10 }}>Deletes all transactions for this pocket only. Other pocket is kept.</div>
              <button style={{ ...s.submitBtn, background:"#E8C547", color:"#1a1a1a", marginTop:0 }}
                onClick={()=>{
                  if(window.confirm(`Delete ALL transactions in ${P.label}? This cannot be undone.`)) {
                    setTxData(prev=>({ ...prev, [pocket]:[] }));
                    setShowSettings(false);
                  }
                }}>
                🗑 Clear {P.label} Transactions
              </button>
            </div>

            <div style={{ height:1, background:"#f0ede8", margin:"20px 0" }}/>

            {/* Full reset */}
            <div style={s.settingsSection}>
              <div style={s.settingsLabel}>Full Reset & Reload Fresh Data</div>
              <div style={{ fontSize:12, color:"#aaa", marginBottom:10 }}>Clears everything and reloads your backed-up data. Use this when switching devices.</div>
              <button style={{ ...s.submitBtn, background:"#E05A5A", marginTop:0 }}
                onClick={()=>{
                  if(window.confirm("This will clear ALL data and reload from backup. Are you sure?")) {
                    localStorage.clear();
                    window.location.reload();
                  }
                }}>
                🔄 Reset & Reload Backup Data
              </button>
            </div>

            <div style={{ height:1, background:"#f0ede8", margin:"20px 0" }}/>

            {/* Change name */}
            <div style={s.settingsSection}>
              <div style={s.settingsLabel}>Change Name / Shop Name</div>
              <div style={{ fontSize:12, color:"#aaa", marginBottom:10 }}>Reset the setup screen to update your name or coffee shop name.</div>
              <button style={{ ...s.submitBtn, background:"#1a1a1a", marginTop:0 }}
                onClick={()=>{
                  if(window.confirm("This will reset your name setup. Data is kept.")) {
                    saveStorage("fin_userName", null);
                    setUserName(null);
                    setShowSettings(false);
                  }
                }}>
                ✏️ Change Name
              </button>
            </div>

            <button style={{ ...s.submitBtn, background:"#f0ede8", color:"#888", marginTop:16 }} onClick={()=>setShowSettings(false)}>
              Close
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Category Editor ──────────────────────────────────────────────────────────
function CategoryEditor({ type, categories, onAdd, onRemove, accent }) {
  const [input, setInput] = useState("");
  const handleAdd = () => { if (input.trim()) { onAdd(input); setInput(""); } };
  return (
    <div style={{ marginBottom:20 }}>
      <div style={{ ...s.sectionTitle, color: type==="income"?"#5aab7e":"#b8962a" }}>
        {type==="income"?"↑ Income Categories":"↓ Expense Categories"}
      </div>
      <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginBottom:10 }}>
        {categories.map(cat=>(
          <div key={cat} style={{ display:"flex", alignItems:"center", gap:4, background:"#f5f5f5", borderRadius:20, padding:"5px 10px 5px 12px", fontSize:13 }}>
            <span>{cat}</span>
            {cat!=="Other" && (
              <button onClick={()=>onRemove(cat)}
                style={{ background:"none", border:"none", cursor:"pointer", color:"#bbb", fontSize:14, lineHeight:1, padding:"0 0 0 2px" }}>×</button>
            )}
          </div>
        ))}
      </div>
      <div style={{ display:"flex", gap:8 }}>
        <input style={{ ...s.input, flex:1, marginBottom:0 }} placeholder={`New ${type} category…`}
          value={input} onChange={e=>setInput(e.target.value)}
          onKeyDown={e=>{ if(e.key==="Enter") handleAdd(); }}/>
        <button style={{ ...s.submitBtn, padding:"10px 16px", marginTop:0, background:accent }} onClick={handleAdd}>Add</button>
      </div>
    </div>
  );
}

// ─── Overview ─────────────────────────────────────────────────────────────────
function Overview({ recent, budgets, spentByCategory, fmt, deleteTransaction, accent }) {
  return (
    <div>
      <div style={s.sectionTitle}>Budget Overview</div>
      {budgets.slice(0,3).map(b=>{
        const spent=spentByCategory(b.category), pct=Math.min((spent/b.limit)*100,100), over=spent>b.limit;
        return (
          <div key={b.category} style={s.budgetRow}>
            <div style={s.budgetTop}>
              <span style={s.budgetCat}>{b.category}</span>
              <span style={{ fontSize:12, color:over?"#E05A5A":"#888" }}>{fmt(spent)} / {fmt(b.limit)}</span>
            </div>
            <div style={s.barTrack}><div style={{ ...s.barFill, width:`${pct}%`, background:over?"#E05A5A":b.color }}/></div>
          </div>
        );
      })}
      <div style={{ ...s.sectionTitle, marginTop:24 }}>Recent</div>
      {recent.map(t=><TxRow key={t.id} t={t} fmt={fmt} onDelete={deleteTransaction} accent={accent}/>)}
    </div>
  );
}

// ─── Transactions ─────────────────────────────────────────────────────────────
function Transactions({ transactions, fmt, deleteTransaction, accent }) {
  const [search,setSearch]=useState("");
  const sorted=[...transactions].sort((a,b)=>new Date(b.date)-new Date(a.date))
    .filter(t=>!search||t.category.toLowerCase().includes(search.toLowerCase())||t.note.toLowerCase().includes(search.toLowerCase()));
  return (
    <div>
      <input style={{ ...s.input, marginBottom:16 }} placeholder="Search…" value={search} onChange={e=>setSearch(e.target.value)}/>
      <div style={s.sectionTitle}>{sorted.length} Transactions</div>
      {sorted.map(t=><TxRow key={t.id} t={t} fmt={fmt} onDelete={deleteTransaction} accent={accent}/>)}
      {sorted.length===0&&<div style={s.empty}>No transactions found.</div>}
    </div>
  );
}

// ─── TxRow ────────────────────────────────────────────────────────────────────
function TxRow({ t, fmt, onDelete }) {
  const [swiped,setSwiped]=useState(false);
  return (
    <div style={{ position:"relative", marginBottom:8, overflow:"hidden", borderRadius:14 }}>
      {swiped&&<button style={s.deleteBtn} onClick={()=>onDelete(t.id)}>Delete</button>}
      <div style={{ ...s.txRow, transform:swiped?"translateX(-80px)":"none", transition:"transform .2s" }} onClick={()=>setSwiped(x=>!x)}>
        <div style={{ ...s.txIcon, background:t.type==="income"?"#7EC8A420":"#f0ede820" }}>{t.type==="income"?"↑":"↓"}</div>
        <div style={{ flex:1 }}>
          <div style={s.txCat}>{t.category}</div>
          <div style={s.txNote}>{t.note||t.date}</div>
        </div>
        <div style={{ textAlign:"right" }}>
          <div style={{ ...s.txAmt, color:t.type==="income"?"#5aab7e":"#1a1a1a" }}>{t.type==="income"?"+":"-"}{fmt(t.amount)}</div>
          <div style={{ fontSize:11, color:"#ccc", marginTop:2 }}>{t.date}</div>
        </div>
      </div>
    </div>
  );
}

// ─── Budgets ──────────────────────────────────────────────────────────────────
function Budgets({ budgets, setBudgets, spentByCategory, fmt, accent, pocketCats }) {
  const [editing,  setEditing]  = useState(null);
  const [newLimit, setNewLimit] = useState("");
  const [showAdd,  setShowAdd]  = useState(false);
  const [newCat,   setNewCat]   = useState("");
  const [newBudget,setNewBudget]= useState("");

  const allExpenseCats = pocketCats.expense;
  const usedCats = budgets.map(b => b.category);
  const availableCats = allExpenseCats.filter(c => !usedCats.includes(c));

  const save = () => {
    if (!newLimit || isNaN(newLimit)) return;
    setBudgets(budgets.map(b => b.category === editing ? { ...b, limit: parseFloat(newLimit) } : b));
    setEditing(null); setNewLimit("");
  };

  const addBudget = () => {
    if (!newCat || !newBudget || isNaN(newBudget)) return;
    const color = COLORS[budgets.length % COLORS.length];
    setBudgets([...budgets, { category: newCat, limit: parseFloat(newBudget), color }]);
    setNewCat(""); setNewBudget(""); setShowAdd(false);
  };

  const deleteBudget = (cat) => {
    setBudgets(budgets.filter(b => b.category !== cat));
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
        <div style={s.sectionTitle}>Monthly Budgets</div>
        <button style={{ ...s.editBtn, fontSize:13, color: accent, border:`1px solid ${accent}`, borderRadius:8, padding:"4px 10px" }}
          onClick={() => setShowAdd(v => !v)}>
          {showAdd ? "Cancel" : "+ Add Budget"}
        </button>
      </div>

      {/* Add new budget form */}
      {showAdd && (
        <div style={{ ...s.budgetCard, background:"#fafafa", marginBottom:16 }}>
          <div style={{ fontSize:13, fontWeight:600, color:"#1a1a1a", marginBottom:10 }}>New Budget</div>
          <select style={{ ...s.input, marginBottom:8 }} value={newCat} onChange={e => setNewCat(e.target.value)}>
            <option value="">Select category…</option>
            {availableCats.map(c => <option key={c}>{c}</option>)}
          </select>
          <input style={{ ...s.input, marginBottom:8 }} type="number" placeholder="Monthly limit ($)"
            value={newBudget} onChange={e => setNewBudget(e.target.value)} />
          <button style={{ ...s.submitBtn, background: newCat && newBudget ? accent : "#ccc", marginTop:0 }}
            disabled={!newCat || !newBudget} onClick={addBudget}>
            Add Budget
          </button>
        </div>
      )}

      {budgets.length === 0 && (
        <div style={s.empty}>No budgets yet. Tap "+ Add Budget" to get started.</div>
      )}

      {budgets.map(b => {
        const spent = spentByCategory(b.category), pct = Math.min((spent/b.limit)*100, 100), over = spent > b.limit;
        return (
          <div key={b.category} style={s.budgetCard}>
            <div style={s.budgetTop}>
              <div>
                <div style={s.budgetCat}>{b.category}</div>
                <div style={{ fontSize:12, color:"#888", marginTop:2 }}>{fmt(Math.max(b.limit-spent,0))} left</div>
              </div>
              <div style={{ textAlign:"right" }}>
                <div style={{ ...s.txAmt, color:over?"#E05A5A":"#1a1a1a" }}>{fmt(spent)}<span style={{ fontSize:11, fontWeight:400, color:"#bbb" }}> / {fmt(b.limit)}</span></div>
                <div style={{ display:"flex", gap:8, justifyContent:"flex-end", marginTop:4 }}>
                  <button style={s.editBtn} onClick={()=>{ setEditing(b.category); setNewLimit(String(b.limit)); }}>Edit</button>
                  <button style={{ ...s.editBtn, color:"#E05A5A" }} onClick={()=>deleteBudget(b.category)}>Delete</button>
                </div>
              </div>
            </div>
            <div style={s.barTrack}><div style={{ ...s.barFill, width:`${pct}%`, background:over?"#E05A5A":b.color }}/></div>
            {editing === b.category && (
              <div style={{ display:"flex", gap:8, marginTop:10 }}>
                <input style={{ ...s.input, flex:1, marginBottom:0 }} type="number" placeholder="New limit"
                  value={newLimit} onChange={e=>setNewLimit(e.target.value)}/>
                <button style={{ ...s.submitBtn, padding:"10px 16px", marginTop:0, background:accent }} onClick={save}>Save</button>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─── Reports ──────────────────────────────────────────────────────────────────
// ─── Period Picker Modal ─────────────────────────────────────────────────────
function PeriodPicker({ transactions, onDownload, onClose, accent }) {
  const years = [...new Set(transactions.map(t=>new Date(t.date).getFullYear()))].sort((a,b)=>b-a);
  const curMonth = new Date().getMonth();
  const curYear  = new Date().getFullYear();
  const [mode,    setMode]    = useState("month");
  const [selYear, setSelYear] = useState(years[0] || curYear);
  const [selMonth,setSelMonth]= useState(curMonth);

  const handleDownload = () => {
    let filterFn, label;
    if (mode === "month") {
      filterFn = t => { const d=new Date(t.date); return d.getFullYear()===selYear && d.getMonth()===selMonth; };
      label = `${MONTHS[selMonth]}-${selYear}`;
    } else if (mode === "year") {
      filterFn = t => new Date(t.date).getFullYear()===selYear;
      label = `${selYear}`;
    } else {
      filterFn = () => true;
      label = "All-Time";
    }
    onDownload(filterFn, label);
    onClose();
  };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.4)", display:"flex", alignItems:"flex-end", justifyContent:"center", zIndex:200 }}
      onClick={onClose}>
      <div style={{ background:"#fff", borderRadius:"24px 24px 0 0", padding:"28px 24px 48px", width:"100%", maxWidth:430 }}
        onClick={e=>e.stopPropagation()}>
        <div style={{ fontSize:17, fontWeight:700, color:"#1a1a1a", marginBottom:20 }}>Choose Report Period</div>

        {/* Mode selector */}
        <div style={{ display:"flex", gap:6, marginBottom:20 }}>
          {[["month","By Month"],["year","By Year"],["all","All Time"]].map(([v,l])=>(
            <button key={v} style={{ flex:1, padding:"9px 4px", borderRadius:10, border:"none", fontSize:12, fontWeight:600, cursor:"pointer",
              background: mode===v ? accent : "#f0ede8", color: mode===v ? "#fff" : "#888" }}
              onClick={()=>setMode(v)}>{l}</button>
          ))}
        </div>

        {/* Year selector */}
        {(mode==="month"||mode==="year") && (
          <div>
            <div style={{ fontSize:11, color:"#bbb", textTransform:"uppercase", letterSpacing:1, marginBottom:8 }}>Year</div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:16 }}>
              {years.map(y=>(
                <button key={y} style={{ padding:"7px 16px", borderRadius:20, border:"none", fontSize:13, fontWeight:600, cursor:"pointer",
                  background: selYear===y ? accent : "#f0ede8", color: selYear===y ? "#fff" : "#888" }}
                  onClick={()=>setSelYear(y)}>{y}</button>
              ))}
            </div>
          </div>
        )}

        {/* Month selector */}
        {mode==="month" && (
          <div>
            <div style={{ fontSize:11, color:"#bbb", textTransform:"uppercase", letterSpacing:1, marginBottom:8 }}>Month</div>
            <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:20 }}>
              {MONTHS.map((m,i)=>(
                <button key={m} style={{ padding:"7px 10px", borderRadius:10, border:"none", fontSize:12, fontWeight:600, cursor:"pointer",
                  background: selMonth===i ? accent : "#f0ede8", color: selMonth===i ? "#fff" : "#888" }}
                  onClick={()=>setSelMonth(i)}>{m}</button>
              ))}
            </div>
          </div>
        )}

        <button style={{ width:"100%", padding:"15px", background:accent, color:"#fff", border:"none", borderRadius:14, fontSize:16, fontWeight:700, cursor:"pointer" }}
          onClick={handleDownload}>
          ⬇ Download Excel
        </button>
      </div>
    </div>
  );
}

function Reports({ transactions, fmt, accent, pocketLabel }) {
  const [view,setView]=useState("monthly");
  const [selectedYear,setSelectedYear]=useState(new Date().getFullYear());
  const [dlFeedback,setDlFeedback]=useState("");
  const [showPicker,setShowPicker]=useState(false);

  const years=[...new Set(transactions.map(t=>new Date(t.date).getFullYear()))].sort((a,b)=>b-a);
  const safeYear = years.includes(selectedYear) ? selectedYear : (years[0]||2026);

  const monthlyData = MONTHS.map((m,mi)=>{
    const mTxs=transactions.filter(t=>{ const d=new Date(t.date); return d.getFullYear()===safeYear&&d.getMonth()===mi; });
    const income=mTxs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
    const expense=mTxs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
    return { month:m, income, expense, net:income-expense, hasTx:mTxs.length>0 };
  });

  const yearlyData=years.map(y=>{
    const yTxs=transactions.filter(t=>new Date(t.date).getFullYear()===y);
    const income=yTxs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
    const expense=yTxs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
    return { year:y, income, expense, net:income-expense };
  });

  const yIncome=monthlyData.reduce((s,m)=>s+m.income,0);
  const yExpense=monthlyData.reduce((s,m)=>s+m.expense,0);
  const yNet=yIncome-yExpense;

  const catSpend={};
  transactions.filter(t=>t.type==="expense"&&new Date(t.date).getFullYear()===safeYear)
    .forEach(t=>{ catSpend[t.category]=(catSpend[t.category]||0)+t.amount; });
  const catList=Object.entries(catSpend).sort((a,b)=>b[1]-a[1]);
  const maxBar=Math.max(...monthlyData.map(m=>Math.max(m.income,m.expense)),1);

  const triggerDownload = (fn) => {
    fn();
    setDlFeedback("Downloaded! ✓");
    setTimeout(()=>setDlFeedback(""),2500);
  };

  const handleXLSXDownload = (filterFn, label) => {
    exportXLSX(transactions, pocketLabel, label, filterFn);
    setDlFeedback("Downloaded! ✓");
    setTimeout(()=>setDlFeedback(""),2500);
  };

  const buildPDF = (filterFn, label) => {
    const filtered = transactions.filter(filterFn);
    const cashIn  = filtered.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
    const cashOut = filtered.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
    const net     = cashIn - cashOut;
    const txRows  = [...filtered].sort((a,b)=>new Date(a.date)-new Date(b.date))
      .map(t=>`<tr><td>${t.date}</td><td>${t.category}</td><td>${t.note||""}</td><td class="g">${t.type==="income"?"$"+t.amount.toLocaleString():""}</td><td class="r">${t.type==="expense"?"$"+t.amount.toLocaleString():""}</td></tr>`)
      .join("");
    exportPDF(`${pocketLabel} — ${label}`,`
      <h1>${pocketLabel} — ${label}</h1>
      <p style="color:#999;margin-bottom:20px">Generated ${new Date().toLocaleDateString()}</p>
      <div class="cards">
        <div class="card"><div class="cl">Cash In</div><div class="cv g">$${cashIn.toLocaleString()}</div></div>
        <div class="card"><div class="cl">Cash Out</div><div class="cv r">$${cashOut.toLocaleString()}</div></div>
        <div class="card"><div class="cl">Net</div><div class="cv ${net>=0?"n":"nb"}">$${net.toLocaleString()}</div></div>
      </div>
      <h2>Transactions</h2>
      <table><thead><tr><th>Date</th><th>Category</th><th>Note</th><th>Cash In</th><th>Cash Out</th></tr></thead><tbody>${txRows}</tbody></table>`);
  };

  return (
    <div>
      {showPicker && <PeriodPicker transactions={transactions} onDownload={handleXLSXDownload} onClose={()=>setShowPicker(false)} accent={accent}/>}

      {/* View toggle */}
      <div style={{ display:"flex", gap:6, marginBottom:20 }}>
        {["monthly","yearly"].map(v=>(
          <button key={v} style={{ ...s.tabBtn, flex:1, ...(view===v?{ ...s.tabActive, background:accent }:{}) }} onClick={()=>setView(v)}>
            {v.charAt(0).toUpperCase()+v.slice(1)}
          </button>
        ))}
      </div>

      {/* Download feedback */}
      {dlFeedback && <div style={s.dlFeedback}>{dlFeedback}</div>}

      {view==="monthly" && (
        <>
          {/* Year tabs */}
          <div style={{ display:"flex", gap:8, marginBottom:16, overflowX:"auto" }}>
            {years.map(y=>(
              <button key={y} style={{ ...s.yearBtn, ...(safeYear===y?{ background:accent, color:"#fff", border:`1px solid ${accent}` }:{}) }}
                onClick={()=>setSelectedYear(y)}>{y}</button>
            ))}
          </div>

          {/* Download row */}
          <div style={s.dlRow}>
            <button style={s.dlBtn} onClick={()=>setShowPicker(true)}>
              ⬇ Excel (.xlsx)
            </button>
            <button style={s.dlBtn} onClick={()=>triggerDownload(()=>buildPDF(
              t=>new Date(t.date).getFullYear()===safeYear,
              String(safeYear)
            ))}>
              🖨 PDF (Print)
            </button>
          </div>

          {/* Summary cards */}
          <div style={s.reportSummaryRow}>
            <div style={{ ...s.reportSummaryCard, borderTop:"3px solid #7EC8A4" }}>
              <div style={s.reportSumLabel}>Income</div>
              <div style={{ ...s.reportSumVal, color:"#5aab7e" }}>{fmt(yIncome)}</div>
            </div>
            <div style={{ ...s.reportSummaryCard, borderTop:"3px solid #E8C547" }}>
              <div style={s.reportSumLabel}>Spent</div>
              <div style={{ ...s.reportSumVal, color:"#b8962a" }}>{fmt(yExpense)}</div>
            </div>
            <div style={{ ...s.reportSummaryCard, borderTop:`3px solid ${accent}` }}>
              <div style={s.reportSumLabel}>Net</div>
              <div style={{ ...s.reportSumVal, color:yNet>=0?accent:"#E05A5A" }}>{fmt(yNet)}</div>
            </div>
          </div>

          {/* Bar chart */}
          <div style={s.sectionTitle}>Month by Month</div>
          <div style={s.chartWrap}>
            {monthlyData.map((m,i)=>(
              <div key={i} style={s.chartCol}>
                <div style={s.barsWrap}>
                  <div style={{ ...s.bar, height:`${(m.income/maxBar)*90}%`, background:"#7EC8A4", opacity:m.hasTx?1:.15 }}/>
                  <div style={{ ...s.bar, height:`${(m.expense/maxBar)*90}%`, background:"#E8C547", opacity:m.hasTx?1:.15 }}/>
                </div>
                <div style={s.chartLabel}>{m.month}</div>
              </div>
            ))}
          </div>
          <div style={{ display:"flex", gap:16, marginBottom:20, marginTop:4 }}>
            <span style={s.legend}><span style={{ ...s.dot, background:"#7EC8A4" }}/> Income</span>
            <span style={s.legend}><span style={{ ...s.dot, background:"#E8C547" }}/> Expenses</span>
          </div>

          {/* Monthly list */}
          <div style={s.sectionTitle}>Monthly Breakdown</div>
          {monthlyData.filter(m=>m.hasTx).map((m,i)=>(
            <div key={i} style={s.reportRow}>
              <div style={s.reportMonth}>{m.month} {safeYear}</div>
              <div style={s.reportCols}>
                <div><div style={s.reportSmall}>Income</div><div style={{ ...s.reportNum, color:"#5aab7e" }}>+{fmt(m.income)}</div></div>
                <div><div style={s.reportSmall}>Spent</div><div style={{ ...s.reportNum, color:"#b8962a" }}>-{fmt(m.expense)}</div></div>
                <div><div style={s.reportSmall}>Net</div><div style={{ ...s.reportNum, color:m.net>=0?accent:"#E05A5A" }}>{fmt(m.net)}</div></div>
              </div>
            </div>
          ))}

          {/* Category breakdown */}
          <div style={{ ...s.sectionTitle, marginTop:24 }}>Top Expenses {safeYear}</div>
          {catList.map(([cat,amt])=>{
            const pct=(amt/(yExpense||1))*100;
            return (
              <div key={cat} style={{ marginBottom:12 }}>
                <div style={s.budgetTop}>
                  <span style={s.budgetCat}>{cat}</span>
                  <span style={{ fontSize:12, color:"#888" }}>{fmt(amt)} · {pct.toFixed(0)}%</span>
                </div>
                <div style={s.barTrack}><div style={{ ...s.barFill, width:`${pct}%`, background:accent+"99" }}/></div>
              </div>
            );
          })}
        </>
      )}

      {view==="yearly" && (
        <>
          {/* Download row */}
          <div style={s.dlRow}>
            <button style={s.dlBtn} onClick={()=>triggerDownload(()=>exportYearlyXLSX(transactions, pocketLabel))}>
              ⬇ Excel (.xlsx)
            </button>
            <button style={s.dlBtn} onClick={()=>triggerDownload(()=>{
              const yRows=yearlyData.map(y=>`<tr><td><b>${y.year}</b></td><td class="g">$${y.income.toLocaleString()}</td><td class="r">$${y.expense.toLocaleString()}</td><td class="${y.net>=0?"n":"nb"}">$${y.net.toLocaleString()}</td></tr>`).join("");
              exportPDF(`${pocketLabel} Yearly Report`,`
                <h1>${pocketLabel} — Yearly Report</h1>
                <p style="color:#999;margin-bottom:20px">Generated ${new Date().toLocaleDateString()}</p>
                <h2>Year over Year</h2>
                <table><thead><tr><th>Year</th><th>Income</th><th>Expenses</th><th>Net</th></tr></thead><tbody>${yRows}</tbody></table>`);
            })}>🖨 PDF (Print)</button>
          </div>

          <div style={s.sectionTitle}>Year over Year</div>
          {yearlyData.map((y,i)=>(
            <div key={i} style={s.budgetCard}>
              <div style={{ fontSize:18, fontWeight:700, color:accent, marginBottom:14 }}>{y.year}</div>
              <div style={s.reportCols}>
                <div><div style={s.reportSmall}>Income</div><div style={{ ...s.reportNum, color:"#5aab7e" }}>{fmt(y.income)}</div></div>
                <div><div style={s.reportSmall}>Expenses</div><div style={{ ...s.reportNum, color:"#b8962a" }}>{fmt(y.expense)}</div></div>
                <div><div style={s.reportSmall}>Net</div><div style={{ ...s.reportNum, color:y.net>=0?accent:"#E05A5A" }}>{fmt(y.net)}</div></div>
              </div>
              <div style={{ marginTop:14 }}>
                <div style={{ display:"flex", gap:4, height:6 }}>
                  <div style={{ flex:y.income, background:"#7EC8A4", borderRadius:3 }}/>
                  <div style={{ flex:y.expense, background:"#E8C547", borderRadius:3 }}/>
                </div>
              </div>
            </div>
          ))}

          {yearlyData.length>1&&(
            <>
              <div style={{ ...s.sectionTitle, marginTop:24 }}>Growth</div>
              {yearlyData.slice(0,-1).map((y,i)=>{
                const prev=yearlyData[i+1];
                const ig=prev.income?((y.income-prev.income)/prev.income*100).toFixed(1):"—";
                const eg=prev.expense?((y.expense-prev.expense)/prev.expense*100).toFixed(1):"—";
                return (
                  <div key={i} style={{ ...s.budgetCard, background:"#fff" }}>
                    <div style={{ fontSize:13, color:"#888", marginBottom:10 }}>{prev.year} → {y.year}</div>
                    <div style={{ display:"flex", gap:16 }}>
                      <div><div style={s.reportSmall}>Revenue</div><div style={{ fontSize:15, fontWeight:700, color:parseFloat(ig)>=0?"#5aab7e":"#E05A5A" }}>{parseFloat(ig)>=0?"▲":"▼"} {Math.abs(ig)}%</div></div>
                      <div><div style={s.reportSmall}>Expenses</div><div style={{ fontSize:15, fontWeight:700, color:parseFloat(eg)>=0?"#E05A5A":"#5aab7e" }}>{parseFloat(eg)>=0?"▲":"▼"} {Math.abs(eg)}%</div></div>
                    </div>
                  </div>
                );
              })}
            </>
          )}
        </>
      )}
    </div>
  );
}

// ─── Styles ───────────────────────────────────────────────────────────────────
const s = {
  shell:             { minHeight:"100vh", background:"#e8e4df", display:"flex", justifyContent:"center", fontFamily:"'DM Sans', sans-serif" },
  app:               { width:"100%", maxWidth:430, minHeight:"100vh", paddingBottom:48 },
  pocketBar:         { display:"flex", gap:8, padding:"52px 20px 0", marginBottom:4 },
  pocketBtn:         { flex:1, padding:"8px 10px", borderRadius:12, border:"1.5px solid #ddd", background:"#fff", fontSize:13, fontWeight:600, color:"#888", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:5 },
  pocketActive:      { color:"#fff", border:"none" },
  header:            { display:"flex", justifyContent:"space-between", alignItems:"center", padding:"16px 24px" },
  headerLabel:       { fontSize:24, fontWeight:700, color:"#1a1a1a", letterSpacing:-0.5 },
  headerSub:         { fontSize:12, color:"#999", marginTop:2 },
  addBtn:            { width:40, height:40, borderRadius:20, color:"#fff", border:"none", fontSize:22, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" },
  balanceCard:       { margin:"0 20px 16px", background:"#fff", borderRadius:20, padding:"22px 24px 18px", boxShadow:"0 2px 16px rgba(0,0,0,.06)" },
  balanceLabel:      { fontSize:11, color:"#aaa", textTransform:"uppercase", letterSpacing:1.2, marginBottom:4 },
  balanceAmount:     { fontSize:38, fontWeight:700, letterSpacing:-1.5, marginBottom:18 },
  balanceRow:        { display:"flex", alignItems:"center" },
  balanceStat:       { flex:1, display:"flex", flexDirection:"column", gap:4 },
  balanceDivider:    { width:1, height:32, background:"#eee", margin:"0 18px" },
  dot:               { width:8, height:8, borderRadius:4, display:"inline-block", marginRight:4 },
  statLabel:         { fontSize:11, color:"#aaa", textTransform:"uppercase", letterSpacing:.8 },
  statVal:           { fontSize:15, fontWeight:600, color:"#1a1a1a" },
  tabs:              { display:"flex", gap:5, padding:"0 20px 16px" },
  tabBtn:            { flex:1, padding:"8px 2px", borderRadius:10, border:"none", background:"transparent", fontSize:12, fontWeight:500, color:"#999", cursor:"pointer" },
  tabActive:         { color:"#fff" },
  content:           { padding:"0 20px" },
  sectionTitle:      { fontSize:11, fontWeight:700, color:"#bbb", textTransform:"uppercase", letterSpacing:1.2, marginBottom:10 },
  budgetRow:         { marginBottom:14 },
  budgetCard:        { background:"#fff", borderRadius:16, padding:16, marginBottom:10, boxShadow:"0 1px 8px rgba(0,0,0,.05)" },
  budgetTop:         { display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:8 },
  budgetCat:         { fontSize:14, fontWeight:600, color:"#1a1a1a" },
  barTrack:          { height:5, background:"#f0ede8", borderRadius:3, overflow:"hidden" },
  barFill:           { height:"100%", borderRadius:3, transition:"width .5s ease" },
  txRow:             { display:"flex", alignItems:"center", gap:12, background:"#fff", borderRadius:14, padding:"12px 14px", cursor:"pointer" },
  txIcon:            { width:34, height:34, borderRadius:10, display:"flex", alignItems:"center", justifyContent:"center", fontSize:15, fontWeight:600, flexShrink:0 },
  txCat:             { fontSize:13, fontWeight:600, color:"#1a1a1a" },
  txNote:            { fontSize:11, color:"#aaa", marginTop:2 },
  txAmt:             { fontSize:14, fontWeight:700 },
  deleteBtn:         { position:"absolute", right:0, top:0, bottom:0, width:80, background:"#E05A5A", color:"#fff", border:"none", fontSize:12, fontWeight:600, cursor:"pointer", borderRadius:"0 14px 14px 0" },
  overlay:           { position:"fixed", inset:0, background:"rgba(0,0,0,.35)", display:"flex", alignItems:"flex-end", justifyContent:"center", zIndex:100 },
  modal:             { background:"#fff", borderRadius:"24px 24px 0 0", padding:"28px 24px 48px", width:"100%", maxWidth:430 },
  modalTitle:        { fontSize:17, fontWeight:700, color:"#1a1a1a" },
  typeToggle:        { display:"flex", gap:8, marginBottom:14 },
  typeBtn:           { flex:1, padding:"10px", borderRadius:12, border:"1.5px solid #eee", background:"#fff", fontSize:14, fontWeight:600, color:"#aaa", cursor:"pointer" },
  input:             { width:"100%", padding:"13px 14px", borderRadius:12, border:"1.5px solid #eee", fontSize:15, marginBottom:10, outline:"none", background:"#fafafa", boxSizing:"border-box", color:"#1a1a1a", fontFamily:"inherit" },
  submitBtn:         { width:"100%", padding:"15px", color:"#fff", border:"none", borderRadius:14, fontSize:16, fontWeight:700, cursor:"pointer", marginTop:4, fontFamily:"inherit" },
  editBtn:           { fontSize:11, color:"#aaa", background:"none", border:"none", cursor:"pointer", padding:0, marginTop:4, fontFamily:"inherit" },
  manageBtn:         { fontSize:12, color:"#aaa", background:"none", border:"1px solid #eee", borderRadius:8, cursor:"pointer", padding:"5px 10px", fontFamily:"inherit" },
  empty:             { textAlign:"center", color:"#bbb", marginTop:40, fontSize:14 },
  reportSummaryRow:  { display:"flex", gap:8, marginBottom:20 },
  reportSummaryCard: { flex:1, background:"#fff", borderRadius:14, padding:"12px 10px", boxShadow:"0 1px 6px rgba(0,0,0,.05)" },
  reportSumLabel:    { fontSize:10, color:"#bbb", textTransform:"uppercase", letterSpacing:.8, marginBottom:4 },
  reportSumVal:      { fontSize:15, fontWeight:700 },
  reportRow:         { background:"#fff", borderRadius:14, padding:"14px 16px", marginBottom:8, boxShadow:"0 1px 6px rgba(0,0,0,.04)" },
  reportMonth:       { fontSize:13, fontWeight:700, color:"#1a1a1a", marginBottom:10 },
  reportCols:        { display:"flex", gap:12 },
  reportSmall:       { fontSize:10, color:"#bbb", textTransform:"uppercase", letterSpacing:.6, marginBottom:3 },
  reportNum:         { fontSize:14, fontWeight:700 },
  chartWrap:         { display:"flex", gap:4, height:100, alignItems:"flex-end", marginBottom:8, background:"#fff", borderRadius:14, padding:"12px 10px 0" },
  chartCol:          { flex:1, display:"flex", flexDirection:"column", alignItems:"center", height:"100%" },
  barsWrap:          { flex:1, width:"100%", display:"flex", gap:1, alignItems:"flex-end" },
  bar:               { flex:1, borderRadius:"3px 3px 0 0", minHeight:2, transition:"height .4s ease" },
  chartLabel:        { fontSize:9, color:"#bbb", marginTop:4, textAlign:"center" },
  legend:            { display:"flex", alignItems:"center", gap:4, fontSize:11, color:"#888" },
  yearBtn:           { padding:"6px 14px", borderRadius:20, border:"1px solid #ddd", background:"#fff", fontSize:13, fontWeight:600, color:"#888", cursor:"pointer", whiteSpace:"nowrap" },
  dlRow:             { display:"flex", gap:8, marginBottom:20 },
  settingsSection:   { marginBottom:4 },
  settingsLabel:     { fontSize:14, fontWeight:700, color:"#1a1a1a", marginBottom:6 },
  dlBtn:             { flex:1, padding:"11px 8px", borderRadius:12, border:"1.5px solid #e0ddd8", background:"#fff", fontSize:13, fontWeight:600, color:"#444", cursor:"pointer", fontFamily:"inherit" },
  dlFeedback:        { background:"#7EC8A4", color:"#fff", borderRadius:10, padding:"9px 14px", marginBottom:14, fontSize:13, fontWeight:600, textAlign:"center" },
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
  * { box-sizing:border-box; margin:0; padding:0; -webkit-tap-highlight-color:transparent; }
  .mounted .fade-in { animation:fadeUp .4s ease both; }
  @keyframes fadeUp { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:none; } }
  input[type=number]::-webkit-inner-spin-button { -webkit-appearance:none; }
  select, input, button { font-family:'DM Sans',sans-serif; }
  select { appearance:none; }
  ::-webkit-scrollbar { display:none; }
`;
