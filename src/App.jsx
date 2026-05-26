import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────
const DEFAULT_CATS = {
  personal: {
    income: ["Salary", "Freelance", "Investment", "Gift", "Other"],
    expense: ["Food", "Transport", "Housing", "Health", "Shopping", "Entertainment", "Other"],
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

// ─── Seed Data ────────────────────────────────────────────────────────────────
const SEED = {
  personal: [
    { id:1,  type:"income",  category:"Salary",        amount:4200, note:"May paycheck",    date:"2026-05-01" },
    { id:2,  type:"expense", category:"Housing",        amount:1400, note:"Rent",            date:"2026-05-02" },
    { id:3,  type:"expense", category:"Food",           amount:85,   note:"Groceries",       date:"2026-05-10" },
    { id:4,  type:"expense", category:"Transport",      amount:45,   note:"Gas",             date:"2026-05-12" },
    { id:5,  type:"income",  category:"Freelance",      amount:800,  note:"Design project",  date:"2026-05-15" },
    { id:6,  type:"expense", category:"Entertainment",  amount:60,   note:"Concert",         date:"2026-05-18" },
    { id:7,  type:"expense", category:"Shopping",       amount:120,  note:"Clothes",         date:"2026-05-20" },
    { id:8,  type:"income",  category:"Salary",         amount:4200, note:"Apr paycheck",    date:"2026-04-01" },
    { id:9,  type:"expense", category:"Housing",        amount:1400, note:"Rent",            date:"2026-04-02" },
    { id:10, type:"expense", category:"Food",           amount:210,  note:"Groceries",       date:"2026-04-15" },
    { id:11, type:"income",  category:"Salary",         amount:4200, note:"Mar paycheck",    date:"2026-03-01" },
    { id:12, type:"expense", category:"Housing",        amount:1400, note:"Rent",            date:"2026-03-02" },
    { id:13, type:"expense", category:"Transport",      amount:90,   note:"Insurance",       date:"2026-03-10" },
    { id:14, type:"income",  category:"Salary",         amount:4200, note:"Feb paycheck",    date:"2026-02-01" },
    { id:15, type:"expense", category:"Housing",        amount:1400, note:"Rent",            date:"2026-02-02" },
    { id:16, type:"income",  category:"Salary",         amount:4200, note:"Jan paycheck",    date:"2026-01-01" },
    { id:17, type:"expense", category:"Shopping",       amount:350,  note:"New Year gear",   date:"2026-01-05" },
  ],
  coffeeshop: [
    { id:101, type:"income",  category:"Coffee Sales",     amount:6800, note:"May revenue",       date:"2026-05-31" },
    { id:102, type:"income",  category:"Food Sales",       amount:1200, note:"Pastries May",      date:"2026-05-31" },
    { id:103, type:"expense", category:"Beans & Supplies", amount:720,  note:"Specialty beans",   date:"2026-05-05" },
    { id:104, type:"expense", category:"Staff Wages",      amount:2800, note:"3 baristas",        date:"2026-05-30" },
    { id:105, type:"expense", category:"Rent",             amount:2000, note:"May rent",          date:"2026-05-01" },
    { id:106, type:"expense", category:"Utilities",        amount:380,  note:"Electricity+water", date:"2026-05-15" },
    { id:107, type:"income",  category:"Coffee Sales",     amount:6400, note:"Apr revenue",       date:"2026-04-30" },
    { id:108, type:"expense", category:"Staff Wages",      amount:2800, note:"Apr wages",         date:"2026-04-30" },
    { id:109, type:"expense", category:"Rent",             amount:2000, note:"Apr rent",          date:"2026-04-01" },
    { id:110, type:"expense", category:"Beans & Supplies", amount:680,  note:"Restock",           date:"2026-04-10" },
    { id:111, type:"income",  category:"Coffee Sales",     amount:5900, note:"Mar revenue",       date:"2026-03-31" },
    { id:112, type:"expense", category:"Staff Wages",      amount:2800, note:"Mar wages",         date:"2026-03-31" },
    { id:113, type:"expense", category:"Rent",             amount:2000, note:"Mar rent",          date:"2026-03-01" },
    { id:114, type:"income",  category:"Coffee Sales",     amount:5200, note:"Feb revenue",       date:"2026-02-28" },
    { id:115, type:"expense", category:"Staff Wages",      amount:2800, note:"Feb wages",         date:"2026-02-28" },
    { id:116, type:"expense", category:"Rent",             amount:2000, note:"Feb rent",          date:"2026-02-01" },
    { id:117, type:"income",  category:"Coffee Sales",     amount:4800, note:"Jan revenue",       date:"2026-01-31" },
    { id:118, type:"expense", category:"Equipment",        amount:1200, note:"Espresso machine",  date:"2026-01-10" },
    { id:119, type:"expense", category:"Rent",             amount:2000, note:"Jan rent",          date:"2026-01-01" },
  ],
};

// ─── Excel Export ─────────────────────────────────────────────────────────────
function exportMonthlyXLSX(transactions, selectedYear, pocketLabel) {
  const MONTHS_FULL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const wb = XLSX.utils.book_new();

  // ── Sheet 1: Monthly Summary ──
  const monthlyRows = [["Month", "Income ($)", "Expenses ($)", "Net ($)"]];
  let totIncome = 0, totExpense = 0;
  MONTHS.forEach((m, mi) => {
    const mTxs = transactions.filter(t => {
      const d = new Date(t.date);
      return d.getFullYear() === selectedYear && d.getMonth() === mi;
    });
    const inc  = mTxs.filter(t => t.type === "income").reduce((s,t) => s+t.amount, 0);
    const exp  = mTxs.filter(t => t.type === "expense").reduce((s,t) => s+t.amount, 0);
    if (inc || exp) { monthlyRows.push([`${m} ${selectedYear}`, inc, exp, inc - exp]); totIncome += inc; totExpense += exp; }
  });
  monthlyRows.push(["TOTAL", totIncome, totExpense, totIncome - totExpense]);
  const ws1 = XLSX.utils.aoa_to_sheet(monthlyRows);
  ws1["!cols"] = [{ wch: 18 },{ wch: 14 },{ wch: 14 },{ wch: 14 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Monthly Summary");

  // ── Sheet 2: All Transactions ──
  const txRows = [["Date","Type","Category","Amount ($)","Note"]];
  [...transactions]
    .filter(t => new Date(t.date).getFullYear() === selectedYear)
    .sort((a,b) => new Date(b.date)-new Date(a.date))
    .forEach(t => txRows.push([t.date, t.type, t.category, t.amount, t.note]));
  const ws2 = XLSX.utils.aoa_to_sheet(txRows);
  ws2["!cols"] = [{ wch: 12 },{ wch: 10 },{ wch: 20 },{ wch: 14 },{ wch: 30 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Transactions");

  // ── Sheet 3: Category Breakdown ──
  const catSpend = {};
  transactions.filter(t => t.type==="expense" && new Date(t.date).getFullYear()===selectedYear)
    .forEach(t => { catSpend[t.category] = (catSpend[t.category]||0)+t.amount; });
  const catRows = [["Category","Amount ($)","% of Total"]];
  Object.entries(catSpend).sort((a,b)=>b[1]-a[1])
    .forEach(([cat,amt]) => catRows.push([cat, amt, `${((amt/(totExpense||1))*100).toFixed(1)}%`]));
  const ws3 = XLSX.utils.aoa_to_sheet(catRows);
  ws3["!cols"] = [{ wch: 22 },{ wch: 14 },{ wch: 14 }];
  XLSX.utils.book_append_sheet(wb, ws3, "Category Breakdown");

  XLSX.writeFile(wb, `${pocketLabel}-${selectedYear}-monthly-report.xlsx`);
}

function exportYearlyXLSX(transactions, pocketLabel) {
  const wb = XLSX.utils.book_new();
  const years = [...new Set(transactions.map(t => new Date(t.date).getFullYear()))].sort((a,b)=>a-b);

  // ── Sheet 1: Year over Year ──
  const yearRows = [["Year","Income ($)","Expenses ($)","Net ($)"]];
  years.forEach(y => {
    const yTxs = transactions.filter(t => new Date(t.date).getFullYear()===y);
    const inc  = yTxs.filter(t=>t.type==="income").reduce((s,t)=>s+t.amount,0);
    const exp  = yTxs.filter(t=>t.type==="expense").reduce((s,t)=>s+t.amount,0);
    yearRows.push([y, inc, exp, inc-exp]);
  });
  const ws1 = XLSX.utils.aoa_to_sheet(yearRows);
  ws1["!cols"] = [{ wch:8 },{ wch:14 },{ wch:14 },{ wch:14 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Year over Year");

  // ── Sheet 2: Growth ──
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
  const [showAdd,   setShowAdd]  = useState(false);
  const [showCats,  setShowCats] = useState(false);
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
    setTxData(prev=>({ ...prev, [pocket]:[{ id:Date.now(), ...form, amount:parseFloat(form.amount) }, ...prev[pocket]] }));
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
          <button style={{ ...s.addBtn, background:P.accent }} onClick={()=>setShowAdd(true)}>+</button>
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
          {tab==="budgets"      && <Budgets budgets={pBudgets} setBudgets={b=>setBudgets(prev=>({ ...prev, [pocket]:b }))} spentByCategory={spentByCategory} fmt={fmt} accent={P.accent}/>}
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
function Budgets({ budgets, setBudgets, spentByCategory, fmt, accent }) {
  const [editing,setEditing]=useState(null);
  const [newLimit,setNewLimit]=useState("");
  const save=()=>{ if(!newLimit||isNaN(newLimit))return; setBudgets(budgets.map(b=>b.category===editing?{ ...b, limit:parseFloat(newLimit) }:b)); setEditing(null); setNewLimit(""); };
  return (
    <div>
      <div style={s.sectionTitle}>Monthly Budgets</div>
      {budgets.map(b=>{
        const spent=spentByCategory(b.category), pct=Math.min((spent/b.limit)*100,100), over=spent>b.limit;
        return (
          <div key={b.category} style={s.budgetCard}>
            <div style={s.budgetTop}>
              <div><div style={s.budgetCat}>{b.category}</div><div style={{ fontSize:12, color:"#888", marginTop:2 }}>{fmt(Math.max(b.limit-spent,0))} left</div></div>
              <div style={{ textAlign:"right" }}>
                <div style={{ ...s.txAmt, color:over?"#E05A5A":"#1a1a1a" }}>{fmt(spent)}<span style={{ fontSize:11, fontWeight:400, color:"#bbb" }}> / {fmt(b.limit)}</span></div>
                <button style={s.editBtn} onClick={()=>{ setEditing(b.category); setNewLimit(String(b.limit)); }}>Edit limit</button>
              </div>
            </div>
            <div style={s.barTrack}><div style={{ ...s.barFill, width:`${pct}%`, background:over?"#E05A5A":b.color }}/></div>
            {editing===b.category&&(
              <div style={{ display:"flex", gap:8, marginTop:10 }}>
                <input style={{ ...s.input, flex:1, marginBottom:0 }} type="number" placeholder="New limit" value={newLimit} onChange={e=>setNewLimit(e.target.value)}/>
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
function Reports({ transactions, fmt, accent, pocketLabel }) {
  const [view,setView]=useState("monthly");
  const [selectedYear,setSelectedYear]=useState(2026);
  const [dlFeedback,setDlFeedback]=useState("");

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

  const buildMonthlyPDF = () => {
    const rows = monthlyData.filter(m=>m.hasTx).map(m=>
      `<tr><td>${m.month} ${safeYear}</td><td class="g">$${m.income.toLocaleString()}</td><td class="r">$${m.expense.toLocaleString()}</td><td class="${m.net>=0?"n":"nb"}">$${m.net.toLocaleString()}</td></tr>`
    ).join("");
    const catRows = catList.map(([cat,amt])=>
      `<tr><td>${cat}</td><td class="r">$${amt.toLocaleString()}</td><td>${((amt/(yExpense||1))*100).toFixed(1)}%</td></tr>`
    ).join("");
    exportPDF(`${pocketLabel} ${safeYear} Monthly Report`,`
      <h1>${pocketLabel} — Monthly Report</h1>
      <p style="color:#999;margin-bottom:20px">Year: ${safeYear} · Generated ${new Date().toLocaleDateString()}</p>
      <div class="cards">
        <div class="card"><div class="cl">Total Income</div><div class="cv g">$${yIncome.toLocaleString()}</div></div>
        <div class="card"><div class="cl">Total Expenses</div><div class="cv r">$${yExpense.toLocaleString()}</div></div>
        <div class="card"><div class="cl">Net Saved</div><div class="cv ${yNet>=0?"n":"nb"}">$${yNet.toLocaleString()}</div></div>
      </div>
      <h2>Month by Month</h2>
      <table><thead><tr><th>Month</th><th>Income</th><th>Expenses</th><th>Net</th></tr></thead><tbody>${rows}</tbody></table>
      <h2>Expense Breakdown</h2>
      <table><thead><tr><th>Category</th><th>Amount</th><th>% of Total</th></tr></thead><tbody>${catRows}</tbody></table>`);
  };

  return (
    <div>
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
            <button style={s.dlBtn} onClick={()=>triggerDownload(()=>exportMonthlyXLSX(transactions, safeYear, pocketLabel))}>
              ⬇ Excel (.xlsx)
            </button>
            <button style={s.dlBtn} onClick={()=>triggerDownload(buildMonthlyPDF)}>
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
