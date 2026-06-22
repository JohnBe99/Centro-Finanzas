import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const T = {
  bg: "#06102A", surface: "#0A1A3E", card: "#0F2152", border: "#1A3470",
  accent: "#C9A84C", accentLight: "#F5C842", mint: "#2DD4A0", red: "#F4526A",
  muted: "#5A7A99", text: "#E8EEF8", textDim: "#8AACC8",
};

// ── HELPERS ──────────────────────────────────────────────────────────────────
// Store raw numbers, display formatted
const fmtDisplay = (n) => {
  if (!n && n !== 0) return "";
  return Math.round(n).toLocaleString("es-CO");
};
const fmtCOP = (n) => "$" + fmtDisplay(n);

// Parse input string → raw number (strips dots/commas)
const parseRaw = (str) => {
  if (!str && str !== 0) return 0;
  const cleaned = String(str).replace(/\./g, "").replace(/,/g, "").replace(/\$/g, "").trim();
  return parseFloat(cleaned) || 0;
};

const today = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
};
const CATS = ["🍽️ Comida","🚌 Transporte","🏠 Hogar","💊 Salud","🎮 Ocio","👗 Ropa","💡 Servicios","📦 Otro"];
const MONTHS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

// ── LOGO ─────────────────────────────────────────────────────────────────────
const LogoSVG = ({ size = 40 }) => (
  <svg width={size} height={size} viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="gold" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" stopColor="#F5C842"/>
        <stop offset="50%" stopColor="#C9A84C"/>
        <stop offset="100%" stopColor="#F5C842"/>
      </linearGradient>
    </defs>
    <path d="M50 8 C50 8 44 18 44 24 C44 28 46.5 30 50 30 C53.5 30 56 28 56 24 C56 18 50 8 50 8Z" fill="url(#gold)"/>
    <path d="M42 22 C36 20 30 26 32 33 C34 38 40 38 44 34 C46 31 45 25 42 22Z" fill="url(#gold)" opacity="0.9"/>
    <path d="M58 22 C64 20 70 26 68 33 C66 38 60 38 56 34 C54 31 55 25 58 22Z" fill="url(#gold)" opacity="0.9"/>
    <path d="M34 28 C26 24 18 32 22 40 C25 46 33 46 38 41 C41 37 39 31 34 28Z" fill="url(#gold)" opacity="0.75"/>
    <path d="M66 28 C74 24 82 32 78 40 C75 46 67 46 62 41 C59 37 61 31 66 28Z" fill="url(#gold)" opacity="0.75"/>
    <path d="M18 58 C24 52 32 50 40 53" stroke="url(#gold)" strokeWidth="3.5" strokeLinecap="round" fill="none"/>
    <path d="M14 66 C22 58 34 55 44 59" stroke="url(#gold)" strokeWidth="3.5" strokeLinecap="round" fill="none" opacity="0.8"/>
    <path d="M82 58 C76 52 68 50 60 53" stroke="url(#gold)" strokeWidth="3.5" strokeLinecap="round" fill="none"/>
    <path d="M86 66 C78 58 66 55 56 59" stroke="url(#gold)" strokeWidth="3.5" strokeLinecap="round" fill="none" opacity="0.8"/>
    <path d="M50 30 L50 62" stroke="url(#gold)" strokeWidth="3" strokeLinecap="round" fill="none"/>
    <path d="M22 72 L78 72" stroke="url(#gold)" strokeWidth="3.5" strokeLinecap="round" fill="none"/>
    <path d="M30 78 L70 78" stroke="url(#gold)" strokeWidth="2.5" strokeLinecap="round" fill="none" opacity="0.6"/>
  </svg>
);

// ── STRATEGIES ────────────────────────────────────────────────────────────────
const STRATEGIES = {
  conservador: {
    label: "🟢 Conservador", debtPct: 0.40, emergencyPct: 0.35, personalPct: 0.25, primaPct: 0.50,
    description: "Equilibrio entre pagar deudas y mantener calidad de vida. Ideal si tienes gastos variables o dependientes.",
    why: [
      { label: "40% a deuda extra", reason: "Avanzas sin asfixiarte. Reduces el saldo y pagas menos intereses, pero con margen para imprevistos." },
      { label: "35% a emergencias", reason: "Construyes un colchón sólido rápido. Así no necesitas volver a usar las tarjetas si algo pasa." },
      { label: "25% personal", reason: "Calidad de vida mínima. Necesitas motivación para sostener el plan a largo plazo." },
      { label: "50% de prima a deuda", reason: "Abonas la mitad del bono semestral. La otra mitad va a emergencias y gastos previstos." },
    ]
  },
  balanceado: {
    label: "🟡 Balanceado", debtPct: 0.60, emergencyPct: 0.25, personalPct: 0.15, primaPct: 0.75,
    description: "El punto óptimo entre velocidad y sostenibilidad. Recomendado para la mayoría de casos.",
    why: [
      { label: "60% a deuda extra", reason: "Mayor parte del sobrante ataca el capital. Cada peso extra que abonás hoy evita que los intereses sigan creciendo mañana." },
      { label: "25% a emergencias", reason: "Construís fondo en paralelo. En 4-5 meses tenés 1 mes de gastos cubierto sin tocar las tarjetas." },
      { label: "15% personal", reason: "Margen mínimo para salidas y gustos. Suficiente para no sentir que vivís como monje." },
      { label: "75% de prima a deuda", reason: "Los bonos son tu arma más poderosa. El 75% puede liquidar una deuda completa de una sola vez." },
    ]
  },
  agresivo: {
    label: "🔴 Agresivo", debtPct: 0.80, emergencyPct: 0.15, personalPct: 0.05, primaPct: 0.95,
    description: "Máxima velocidad de salida. Para quienes prefieren sacrificar comodidad hoy por libertad rápida.",
    why: [
      { label: "80% a deuda extra", reason: "Casi todo el sobrante va al capital. Los intereses son literalmente dinero que se quema: entre más rápido bajas el saldo, menos pagas en total." },
      { label: "15% a emergencias", reason: "Fondo mínimo de seguridad. No lo toques salvo emergencia real — si lo usás, reponelo antes del siguiente mes." },
      { label: "5% personal", reason: "Solo lo esencial. Este modo es temporal y requiere disciplina fuerte. No es para siempre." },
      { label: "95% de prima a deuda", reason: "Prácticamente toda la prima va directo al capital. Esto puede acortar el plan en meses enteros." },
    ]
  }
};

// ── CLAUDE API ────────────────────────────────────────────────────────────────
async function askClaude(systemPrompt, userMessage) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": process.env.REACT_APP_ANTHROPIC_KEY,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6", max_tokens: 1000,
      system: systemPrompt,
      messages: [{ role: "user", content: userMessage }],
    }),
  });
  const data = await res.json();
  return data.content?.map(b => b.text || "").join("") || "Sin respuesta.";
}

// ── UI COMPONENTS ─────────────────────────────────────────────────────────────
const Card = ({ children, style = {} }) => (
  <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 16, padding: 20, ...style }}>
    {children}
  </div>
);

const Badge = ({ children, color = T.mint }) => (
  <span style={{ background: color+"22", color, border: `1px solid ${color}44`, borderRadius: 20, padding: "3px 10px", fontSize: 11, fontWeight: 700 }}>{children}</span>
);

const Btn = ({ children, onClick, variant = "primary", disabled, style = {} }) => {
  const variants = {
    primary: { background: `linear-gradient(135deg, ${T.accentLight}, ${T.accent})`, color: "#06102A" },
    ghost: { background: T.border+"88", color: T.text, border: "none" },
    danger: { background: T.red+"22", color: T.red, border: `1px solid ${T.red}44` },
  };
  return (
    <button onClick={onClick} disabled={disabled} style={{
      borderRadius: 10, padding: "10px 18px", fontFamily: "inherit",
      fontWeight: 700, fontSize: 13, cursor: disabled ? "not-allowed" : "pointer",
      opacity: disabled ? 0.5 : 1, transition: "all .15s", border: "none",
      ...variants[variant], ...style
    }}>{children}</button>
  );
};

// ── MONEY INPUT — shows formatted, stores raw number ─────────────────────────
const MoneyInput = ({ label, value, onChange }) => {
  // value is a raw number (or ""); display is formatted string
  const [display, setDisplay] = useState(value !== "" ? fmtDisplay(value) : "");

  useEffect(() => {
    if (value === "" || value === 0) setDisplay("");
    else setDisplay(fmtDisplay(value));
  }, [value]);

  const handleChange = (e) => {
    const raw = e.target.value.replace(/\./g, "").replace(/,/g, "").replace(/[^0-9]/g, "");
    const num = raw === "" ? "" : parseInt(raw, 10);
    setDisplay(raw === "" ? "" : parseInt(raw, 10).toLocaleString("es-CO"));
    onChange(num === "" ? "" : num);
  };

  return (
    <div style={{ marginBottom: 14 }}>
      {label && <div style={{ fontSize: 11, color: T.muted, marginBottom: 5, letterSpacing: 1, textTransform: "uppercase" }}>{label}</div>}
      <div style={{ position: "relative" }}>
        <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: T.accent, fontSize: 13, fontWeight: 700 }}>$</span>
        <input
          type="text" inputMode="numeric" value={display}
          onChange={handleChange}
          style={{ width: "100%", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10,
            padding: "10px 12px 10px 28px", color: T.text, fontFamily: "inherit",
            fontSize: 14, outline: "none", boxSizing: "border-box" }}
        />
      </div>
    </div>
  );
};

// Regular text input (for names, dates, etc.)
const TextInput = ({ label, value, onChange, type = "text" }) => (
  <div style={{ marginBottom: 14 }}>
    {label && <div style={{ fontSize: 11, color: T.muted, marginBottom: 5, letterSpacing: 1, textTransform: "uppercase" }}>{label}</div>}
    <input type={type} value={value} onChange={e => onChange(e.target.value)}
      style={{ width: "100%", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10,
        padding: "10px 12px", color: T.text, fontFamily: "inherit", fontSize: 14, outline: "none", boxSizing: "border-box" }} />
  </div>
);

// Rate input (percentage)
const RateInput = ({ label, value, onChange }) => (
  <div style={{ marginBottom: 14 }}>
    {label && <div style={{ fontSize: 11, color: T.muted, marginBottom: 5, letterSpacing: 1, textTransform: "uppercase" }}>{label}</div>}
    <div style={{ position: "relative" }}>
      <input type="number" value={value} onChange={e => onChange(e.target.value)} min="0" max="200" step="0.01"
        style={{ width: "100%", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10,
          padding: "10px 30px 10px 12px", color: T.text, fontFamily: "inherit", fontSize: 14, outline: "none", boxSizing: "border-box" }} />
      <span style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", color: T.muted, fontSize: 12 }}>%</span>
    </div>
  </div>
);

const Select = ({ label, value, onChange, options }) => (
  <div style={{ marginBottom: 14 }}>
    {label && <div style={{ fontSize: 11, color: T.muted, marginBottom: 5, letterSpacing: 1, textTransform: "uppercase" }}>{label}</div>}
    <select value={value} onChange={e => onChange(e.target.value)}
      style={{ width: "100%", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10,
        padding: "10px 12px", color: T.text, fontFamily: "inherit", fontSize: 14, outline: "none" }}>
      {options.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  </div>
);

const ProgressBar = ({ pct, color = T.mint }) => (
  <div style={{ background: T.border+"66", borderRadius: 6, height: 8, overflow: "hidden" }}>
    <div style={{ width: `${Math.min(Math.max(pct,0),100)}%`, background: color, height: "100%", borderRadius: 6, transition: "width .6s ease" }} />
  </div>
);

const TABS = [
  { id: "setup", label: "📋 Perfil" },
  { id: "plan",  label: "🎯 Plan" },
  { id: "gastos",label: "💸 Gastos" },
  { id: "ia",    label: "🤖 IA" },
  { id: "reporte",label:"📊 Reporte" },
];

// ── MAIN APP ──────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("setup");
  const [strategy, setStrategy] = useState("balanceado");
  const [showWhyModal, setShowWhyModal] = useState(false);

  // All money values stored as raw numbers ("" when empty)
  const [income, setIncome] = useState("");
  const [semestralPrima, setSemestralPrima] = useState("");
  const [fixedExpenses, setFixedExpenses] = useState([
    { name: "Arriendo", amount: "" },
    { name: "Servicios", amount: "" },
    { name: "Comida", amount: "" },
  ]);
  const [debts, setDebts] = useState([
    { name: "Tarjeta Nu", balance: "", rate: "", minPayment: "" },
    { name: "Tarjeta BBVA Aqua", balance: "", rate: "", minPayment: "" },
  ]);
  const [expenses, setExpenses] = useState([]);
  const [newExp, setNewExp] = useState({ date: today(), desc: "", amount: "", cat: CATS[0] });
  const [iaQ, setIaQ] = useState("");
  const [iaHistory, setIaHistory] = useState([]);
  const [iaLoading, setIaLoading] = useState(false);

  // ── COMPUTED ─────────────────────────────────────────────────────────────────
  const strat = STRATEGIES[strategy];
  const incomeN = Number(income) || 0;
  const totalFixed = fixedExpenses.reduce((s, e) => s + (Number(e.amount) || 0), 0);
  const totalDebtMin = debts.reduce((s, d) => s + (Number(d.minPayment) || 0), 0);
  const surplus = incomeN - totalFixed - totalDebtMin;
  const totalDebt = debts.reduce((s, d) => s + (Number(d.balance) || 0), 0);
  const primaVal = Number(semestralPrima) || 0;
  const debtExtra = Math.max(0, surplus * strat.debtPct);
  const emergencyAmt = Math.max(0, surplus * strat.emergencyPct);
  const personalAmt = Math.max(0, surplus * strat.personalPct);
  const sortedDebts = [...debts].sort((a, b) => (Number(b.rate)||0) - (Number(a.rate)||0));
  const priorityDebt = sortedDebts.find(d => Number(d.balance) > 0);

  // ── PLAN PROJECTION ───────────────────────────────────────────────────────────
  const buildPlan = () => {
    if (!incomeN || debts.length === 0 || surplus <= 0) return [];
    const sorted = [...debts].sort((a, b) => (Number(b.rate)||0) - (Number(a.rate)||0));
    let balances = sorted.map(d => Number(d.balance) || 0);
    const rates = sorted.map(d => (Number(d.rate)||0) / 100 / 12);
    const mins = sorted.map(d => Number(d.minPayment) || 0);
    const extra = Math.max(0, surplus * strat.debtPct);
    const months = [];
    let m = 0;
    while (balances.some(b => b > 100) && m < 72) {
      m++;
      const d = new Date(); d.setMonth(d.getMonth() + m);
      const monthLabel = `${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
      let extraLeft = extra;
      if (m % 6 === 0) extraLeft += primaVal * strat.primaPct;
      for (let i = 0; i < balances.length; i++) {
        if (balances[i] <= 0) continue;
        const interest = balances[i] * rates[i];
        const isFirst = i === balances.findIndex(b => b > 0);
        const pay = Math.min(balances[i] + interest, mins[i] + (isFirst ? extraLeft : 0));
        extraLeft = Math.max(0, extraLeft - Math.max(0, pay - mins[i]));
        balances[i] = Math.max(0, balances[i] + interest - pay);
      }
      months.push({ label: monthLabel, balances: [...balances], total: balances.reduce((s,b)=>s+b,0) });
    }
    return months;
  };

  const plan = buildPlan();
  const debtFreeIdx = plan.findIndex(p => p.total <= 100);
  const debtFreeLabel = debtFreeIdx >= 0 ? plan[debtFreeIdx]?.label : null;

  // ── EXPENSE AGGREGATES ────────────────────────────────────────────────────────
  const expByMonth = expenses.reduce((acc, e) => {
    const k = e.date.slice(0,7);
    acc[k] = (acc[k]||0) + (Number(e.amount)||0);
    return acc;
  }, {});
  const expByCat = expenses.reduce((acc, e) => {
    acc[e.cat] = (acc[e.cat]||0) + (Number(e.amount)||0);
    return acc;
  }, {});
  const thisMonth = today().slice(0,7);
  const thisMonthTotal = expByMonth[thisMonth] || 0;

  // ── DYNAMIC RULES ─────────────────────────────────────────────────────────────
  const dynamicRules = () => {
    if (!incomeN) return ["Ingresa tus datos en Perfil para ver tus reglas personalizadas."];
    const s = [...debts].sort((a,b)=>(Number(b.rate)||0)-(Number(a.rate)||0));
    const first = s[0]; const others = s.slice(1);
    const rules = [];
    if (first?.name) rules.push(`No uses ${first.name} para compras nuevas — cada compra suma intereses al saldo que estás pagando`);
    if (others.length > 0) rules.push(`Paga solo el mínimo de ${others.map(d=>d.name).join(" y ")} — todo el extra va a ${first?.name||"la deuda prioritaria"}`);
    rules.push(`Abona ${fmtCOP(debtExtra)} extra a ${first?.name||"tu deuda prioritaria"} cada mes (${Math.round(strat.debtPct*100)}% del sobrante)`);
    rules.push(`Guarda ${fmtCOP(emergencyAmt)} en una cuenta separada de emergencias (${Math.round(strat.emergencyPct*100)}% del sobrante) — no la toques`);
    if (primaVal > 0) rules.push(`Con tu prima de ${fmtCOP(primaVal)}, abona ${fmtCOP(primaVal*strat.primaPct)} directo a ${first?.name||"la deuda prioritaria"}`);
    rules.push(`Ingresos extra: 70% deuda → 20% emergencias → 10% para ti`);
    if (debtFreeLabel) rules.push(`Siguiendo este plan estarás libre de deudas en ${debtFreeLabel} 🏁`);
    return rules;
  };

  // ── IA CONTEXT ────────────────────────────────────────────────────────────────
  const buildContext = () => `
Eres el consejero de "Centro Conciencia Financiera", experto financiero para Latinoamérica (Colombia).
Situación del usuario:
- Ingreso mensual neto: ${fmtCOP(incomeN)}
- Gastos fijos: ${fixedExpenses.map(e=>`${e.name} ${fmtCOP(Number(e.amount)||0)}`).join(", ")}
- Deudas (orden de ataque): ${[...debts].sort((a,b)=>(Number(b.rate)||0)-(Number(a.rate)||0)).map(d=>`${d.name}: ${fmtCOP(Number(d.balance)||0)} al ${d.rate}% E.A., cuota mínima ${fmtCOP(Number(d.minPayment)||0)}`).join("; ")}
- Sobrante mensual: ${fmtCOP(surplus)}
- Estrategia: ${strat.label}
- Distribución sobrante: ${Math.round(strat.debtPct*100)}% deuda (${fmtCOP(debtExtra)}), ${Math.round(strat.emergencyPct*100)}% emergencias (${fmtCOP(emergencyAmt)}), ${Math.round(strat.personalPct*100)}% personal (${fmtCOP(personalAmt)})
- Prima semestral: ${fmtCOP(primaVal)} → ${Math.round(strat.primaPct*100)}% a deuda
- Libre de deudas proyectado: ${debtFreeLabel || "por calcular (faltan datos)"}
- Gastos registrados este mes: ${fmtCOP(thisMonthTotal)}
Responde en español, claro y directo, máximo 4 párrafos, consejos accionables.
`;

  const sendIA = async () => {
    if (!iaQ.trim() || iaLoading) return;
    const q = iaQ; setIaQ("");
    setIaHistory(h => [...h, { role: "user", text: q }]);
    setIaLoading(true);
    try {
      const resp = await askClaude(buildContext(), q);
      setIaHistory(h => [...h, { role: "ai", text: resp }]);
    } catch {
      setIaHistory(h => [...h, { role: "ai", text: "Error de conexión. Intenta de nuevo." }]);
    }
    setIaLoading(false);
  };

  // ── EXPENSE ACTIONS ───────────────────────────────────────────────────────────
  const addExpense = () => {
    if (!newExp.amount || !newExp.desc) return;
    setExpenses(e => [...e, { ...newExp, id: Date.now() }]);
    setNewExp({ date: today(), desc: "", amount: "", cat: CATS[0] });
  };
  const removeExpense = id => setExpenses(e => e.filter(x => x.id !== id));

  // ── EXPORT XLSX ───────────────────────────────────────────────────────────────
  const exportXLSX = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Gastos
    const gastoRows = expenses.map(e => ({
      "Fecha": e.date,
      "Descripción": e.desc,
      "Categoría": e.cat,
      "Monto (COP)": Number(e.amount) || 0,
    }));
    const wsGastos = XLSX.utils.json_to_sheet(gastoRows.length ? gastoRows : [{"Fecha":"","Descripción":"Sin gastos registrados","Categoría":"","Monto (COP)":0}]);
    wsGastos["!cols"] = [{wch:12},{wch:30},{wch:18},{wch:15}];
    XLSX.utils.book_append_sheet(wb, wsGastos, "Gastos");

    // Sheet 2: Resumen financiero
    const resumen = [
      { "Concepto": "Ingreso mensual neto",   "Valor (COP)": incomeN },
      { "Concepto": "Total gastos fijos",      "Valor (COP)": totalFixed },
      { "Concepto": "Total cuotas mínimas",    "Valor (COP)": totalDebtMin },
      { "Concepto": "Sobrante mensual",        "Valor (COP)": surplus },
      { "Concepto": "Deuda total",             "Valor (COP)": totalDebt },
      { "Concepto": "Prima semestral",         "Valor (COP)": primaVal },
      { "Concepto": "Estrategia elegida",      "Valor (COP)": strat.label },
      { "Concepto": "Abono extra mensual",     "Valor (COP)": debtExtra },
      { "Concepto": "Aporte emergencias/mes",  "Valor (COP)": emergencyAmt },
      { "Concepto": "Colchón personal/mes",    "Valor (COP)": personalAmt },
      { "Concepto": "Proyección libre deudas", "Valor (COP)": debtFreeLabel || "Sin datos suficientes" },
    ];
    const wsResumen = XLSX.utils.json_to_sheet(resumen);
    wsResumen["!cols"] = [{wch:28},{wch:22}];
    XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen");

    // Sheet 3: Deudas
    const debtRows = sortedDebts.map((d,i) => ({
      "Prioridad": i+1,
      "Nombre": d.name,
      "Saldo (COP)": Number(d.balance)||0,
      "Tasa E.A. (%)": Number(d.rate)||0,
      "Cuota mínima (COP)": Number(d.minPayment)||0,
    }));
    const wsDeudas = XLSX.utils.json_to_sheet(debtRows.length ? debtRows : [{"Prioridad":"-","Nombre":"Sin deudas","Saldo (COP)":0,"Tasa E.A. (%)":0,"Cuota mínima (COP)":0}]);
    wsDeudas["!cols"] = [{wch:10},{wch:22},{wch:16},{wch:14},{wch:20}];
    XLSX.utils.book_append_sheet(wb, wsDeudas, "Deudas");

    // Sheet 4: Plan proyectado
    if (plan.length > 0) {
      const planRows = plan.map(p => {
        const row = { "Mes": p.label, "Deuda Total (COP)": Math.round(p.total) };
        sortedDebts.forEach((d,i) => { row[d.name+" (COP)"] = Math.round(p.balances[i]||0); });
        return row;
      });
      const wsPlan = XLSX.utils.json_to_sheet(planRows);
      XLSX.utils.book_append_sheet(wb, wsPlan, "Plan Proyectado");
    }

    XLSX.writeFile(wb, "centro-conciencia-financiera.xlsx");
  };

  const chatEndRef = useRef(null);
  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior: "smooth" }); }, [iaHistory]);

  // ── RENDER ────────────────────────────────────────────────────────────────────
  return (
    <div style={{ background: T.bg, minHeight: "100vh", color: T.text, fontFamily: "'Inter',system-ui,sans-serif", fontSize: 14 }}>

      {/* HEADER */}
      <div style={{ background: `linear-gradient(135deg,#06102A,#0A1A3E,#0D2060)`, borderBottom:`1px solid ${T.accent}33`, padding:"14px 20px", display:"flex", alignItems:"center", justifyContent:"space-between", boxShadow:`0 2px 20px rgba(201,168,76,0.1)` }}>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          <LogoSVG size={46} />
          <div>
            <div style={{ fontWeight:800, fontSize:17, letterSpacing:2, textTransform:"uppercase", background:`linear-gradient(135deg,${T.accentLight},${T.accent})`, WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent" }}>CENTRO</div>
            <div style={{ fontSize:9, letterSpacing:2.5, textTransform:"uppercase", color:T.accent, opacity:0.8, marginTop:1 }}>CONCIENCIA FINANCIERA</div>
          </div>
        </div>
        {debtFreeLabel && <Badge color={T.mint}>🏁 Libre en {debtFreeLabel}</Badge>}
      </div>

      {/* STRATEGY SELECTOR */}
      <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"12px 16px" }}>
        <div style={{ fontSize:10, color:T.muted, marginBottom:8, letterSpacing:1.5, textTransform:"uppercase" }}>Estrategia de pago</div>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          {Object.entries(STRATEGIES).map(([key, s]) => {
            const active = strategy === key;
            const colors = { conservador:T.mint, balanceado:T.accentLight, agresivo:T.red };
            return (
              <button key={key} onClick={()=>setStrategy(key)} style={{ flex:1, padding:"8px 4px", borderRadius:10, fontFamily:"inherit", fontWeight:700, fontSize:11, cursor:"pointer", border:active?`1px solid ${colors[key]}`:`1px solid ${T.border}`, background:active?colors[key]+"22":T.card, color:active?colors[key]:T.muted, transition:"all .2s" }}>{s.label}</button>
            );
          })}
          <button onClick={()=>setShowWhyModal(true)} style={{ background:T.accent+"11", border:`1px solid ${T.accent}44`, borderRadius:10, padding:"8px 12px", color:T.accent, cursor:"pointer", fontFamily:"inherit", fontSize:11, fontWeight:700, whiteSpace:"nowrap" }}>¿Por qué? 💡</button>
        </div>
        <div style={{ fontSize:11, color:T.muted, marginTop:8 }}>{strat.description}</div>
      </div>

      {/* TABS */}
      <div style={{ display:"flex", overflowX:"auto", background:T.surface, borderBottom:`1px solid ${T.border}` }}>
        {TABS.map(t => (
          <button key={t.id} onClick={()=>setTab(t.id)} style={{ flex:1, background:"transparent", color:tab===t.id?T.accent:T.muted, border:"none", borderBottom:tab===t.id?`2px solid ${T.accent}`:"2px solid transparent", padding:"12px 8px", fontFamily:"inherit", fontWeight:700, fontSize:11, cursor:"pointer", whiteSpace:"nowrap", transition:"all .15s" }}>{t.label}</button>
        ))}
      </div>

      {/* WHY MODAL */}
      {showWhyModal && (
        <div onClick={()=>setShowWhyModal(false)} style={{ position:"fixed", inset:0, background:"#000000cc", zIndex:100, display:"flex", alignItems:"flex-end" }}>
          <div onClick={e=>e.stopPropagation()} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:"20px 20px 0 0", padding:24, width:"100%", maxWidth:640, margin:"0 auto", maxHeight:"80vh", overflowY:"auto" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
              <div style={{ fontWeight:800, fontSize:16, color:T.accent }}>¿Por qué se distribuye así? 💡</div>
              <button onClick={()=>setShowWhyModal(false)} style={{ background:"none", border:"none", color:T.muted, fontSize:20, cursor:"pointer" }}>✕</button>
            </div>
            <div style={{ background:T.card, borderRadius:12, padding:14, marginBottom:16, borderLeft:`3px solid ${T.accent}` }}>
              <div style={{ fontWeight:700, color:T.accent, marginBottom:4 }}>{strat.label}</div>
              <div style={{ fontSize:13, color:T.textDim, lineHeight:1.6 }}>{strat.description}</div>
            </div>
            {strat.why.map((w,i)=>(
              <div key={i} style={{ background:T.card, borderRadius:12, padding:14, marginBottom:10 }}>
                <div style={{ fontWeight:700, fontSize:13, color:T.accentLight, marginBottom:6 }}>📌 {w.label}</div>
                <div style={{ fontSize:13, color:T.textDim, lineHeight:1.6 }}>{w.reason}</div>
              </div>
            ))}
            <div style={{ background:T.mint+"11", border:`1px solid ${T.mint}33`, borderRadius:12, padding:14, marginTop:8 }}>
              <div style={{ fontWeight:700, color:T.mint, marginBottom:6 }}>🧮 La lógica del método Avalancha</div>
              <div style={{ fontSize:13, color:T.textDim, lineHeight:1.7 }}>Las tarjetas cobran interés sobre el saldo diario. Cada peso extra que abonás hoy reduce el capital sobre el que te cobran mañana. Por eso atacás siempre la deuda con mayor tasa — es como ganar esa tasa de retorno garantizado.</div>
            </div>
            <Btn onClick={()=>setShowWhyModal(false)} style={{ width:"100%", marginTop:16 }}>Entendido</Btn>
          </div>
        </div>
      )}

      <div style={{ padding:"20px 16px", maxWidth:640, margin:"0 auto" }}>

        {/* ── PERFIL ── */}
        {tab === "setup" && (
          <div>
            <div style={{ fontWeight:800, fontSize:18, marginBottom:4 }}>Tu perfil financiero</div>
            <div style={{ color:T.muted, fontSize:13, marginBottom:20 }}>Ingresa tus datos para personalizar tu plan.</div>

            <Card style={{ marginBottom:16, borderColor:T.accent+"44" }}>
              <div style={{ fontWeight:700, marginBottom:14, color:T.accent }}>💰 Ingresos</div>
              <MoneyInput label="Ingreso mensual neto" value={income} onChange={setIncome} />
              <MoneyInput label="Prima semestral estimada" value={semestralPrima} onChange={setSemestralPrima} />
            </Card>

            <Card style={{ marginBottom:16 }}>
              <div style={{ fontWeight:700, marginBottom:14, color:T.accent }}>🏠 Gastos fijos mensuales</div>
              {fixedExpenses.map((e,i)=>(
                <div key={i} style={{ display:"flex", gap:10, alignItems:"flex-end" }}>
                  <div style={{ flex:1 }}>
                    <TextInput label={i===0?"Nombre":""} value={e.name} onChange={v=>setFixedExpenses(fe=>fe.map((x,j)=>j===i?{...x,name:v}:x))} />
                  </div>
                  <div style={{ flex:1 }}>
                    <MoneyInput label={i===0?"Monto":""} value={e.amount} onChange={v=>setFixedExpenses(fe=>fe.map((x,j)=>j===i?{...x,amount:v}:x))} />
                  </div>
                  <Btn variant="danger" onClick={()=>setFixedExpenses(fe=>fe.filter((_,j)=>j!==i))} style={{ marginBottom:14, padding:"10px 12px" }}>✕</Btn>
                </div>
              ))}
              <Btn variant="ghost" onClick={()=>setFixedExpenses(fe=>[...fe,{name:"",amount:""}])}>+ Agregar gasto</Btn>
            </Card>

            <Card style={{ marginBottom:16 }}>
              <div style={{ fontWeight:700, marginBottom:14, color:T.red }}>💳 Deudas</div>
              {debts.map((d,i)=>(
                <div key={i} style={{ background:T.surface, borderRadius:12, padding:14, marginBottom:12, border:`1px solid ${T.border}` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
                    <input value={d.name} onChange={e=>setDebts(db=>db.map((x,j)=>j===i?{...x,name:e.target.value}:x))}
                      style={{ background:"transparent", border:"none", color:T.text, fontWeight:700, fontSize:14, fontFamily:"inherit", outline:"none", flex:1 }} />
                    <Btn variant="danger" onClick={()=>setDebts(db=>db.filter((_,j)=>j!==i))} style={{ padding:"4px 10px", fontSize:11 }}>✕</Btn>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:10 }}>
                    <MoneyInput label="Saldo" value={d.balance} onChange={v=>setDebts(db=>db.map((x,j)=>j===i?{...x,balance:v}:x))} />
                    <RateInput label="Tasa E.A." value={d.rate} onChange={v=>setDebts(db=>db.map((x,j)=>j===i?{...x,rate:v}:x))} />
                    <MoneyInput label="Cuota mín." value={d.minPayment} onChange={v=>setDebts(db=>db.map((x,j)=>j===i?{...x,minPayment:v}:x))} />
                  </div>
                </div>
              ))}
              <Btn variant="ghost" onClick={()=>setDebts(db=>[...db,{name:"Nueva deuda",balance:"",rate:"",minPayment:""}])}>+ Agregar deuda</Btn>
            </Card>

            {incomeN > 0 && (
              <Card style={{ border:`1px solid ${T.accent}44`, background:T.accent+"06" }}>
                <div style={{ fontWeight:700, marginBottom:12, color:T.accent }}>📊 Resumen actual</div>
                {[
                  ["Ingreso", fmtCOP(incomeN), T.mint],
                  ["Gastos fijos", `-${fmtCOP(totalFixed)}`, T.red],
                  ["Cuotas mínimas", `-${fmtCOP(totalDebtMin)}`, T.red],
                  ["Deuda total", fmtCOP(totalDebt), T.accentLight],
                ].map(([l,v,c])=>(
                  <div key={l} style={{ display:"flex", justifyContent:"space-between", padding:"7px 0", borderBottom:`1px solid ${T.border}44` }}>
                    <span style={{ color:T.textDim }}>{l}</span>
                    <span style={{ color:c, fontWeight:700 }}>{v}</span>
                  </div>
                ))}
                <div style={{ display:"flex", justifyContent:"space-between", padding:"10px 0 0" }}>
                  <span style={{ fontWeight:700 }}>Sobrante mensual</span>
                  <span style={{ color:surplus>=0?T.mint:T.red, fontWeight:800, fontSize:18 }}>{fmtCOP(surplus)}</span>
                </div>
              </Card>
            )}
          </div>
        )}

        {/* ── PLAN ── */}
        {tab === "plan" && (
          <div>
            <div style={{ fontWeight:800, fontSize:18, marginBottom:4 }}>Plan de eliminación de deudas</div>
            <div style={{ color:T.muted, fontSize:13, marginBottom:20 }}>Método Avalancha · {strat.label}</div>

            {totalDebt === 0 ? (
              <Card><div style={{ color:T.muted, textAlign:"center", padding:30 }}>Agrega tus deudas en Perfil para ver el plan.</div></Card>
            ) : (
              <>
                {sortedDebts.map((d,i)=>{
                  const bal = Number(d.balance)||0;
                  const pct = totalDebt > 0 ? (bal/totalDebt)*100 : 0;
                  return (
                    <Card key={i} style={{ marginBottom:12 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", marginBottom:8 }}>
                        <div>
                          <div style={{ fontWeight:700 }}>{d.name}</div>
                          <div style={{ fontSize:11, color:T.muted }}>{d.rate}% E.A. · Cuota mín: {fmtCOP(Number(d.minPayment)||0)}</div>
                        </div>
                        <div style={{ textAlign:"right" }}>
                          <div style={{ fontWeight:800, color:T.red, fontSize:16 }}>{fmtCOP(bal)}</div>
                          {i===0 ? <Badge color={T.accentLight}>🔥 Atacar primero</Badge> : <Badge color={T.muted}>Solo mínimo</Badge>}
                        </div>
                      </div>
                      <ProgressBar pct={pct} color={i===0?T.red:T.accent} />
                      <div style={{ fontSize:11, color:T.muted, marginTop:5 }}>{pct.toFixed(0)}% de la deuda total</div>
                    </Card>
                  );
                })}

                <Card style={{ marginBottom:16 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                    <div style={{ fontWeight:700, color:T.accent }}>💡 Distribución del sobrante ({fmtCOP(surplus)})</div>
                    <button onClick={()=>setShowWhyModal(true)} style={{ background:"none", border:`1px solid ${T.accent}44`, borderRadius:20, padding:"4px 10px", color:T.accent, cursor:"pointer", fontFamily:"inherit", fontSize:11 }}>¿Por qué? 💡</button>
                  </div>
                  {[
                    [`Abono extra → ${priorityDebt?.name||"deuda prioritaria"}`, debtExtra, T.red, strat.debtPct],
                    ["Fondo de emergencia", emergencyAmt, T.mint, strat.emergencyPct],
                    ["Colchón personal", personalAmt, T.muted, strat.personalPct],
                  ].map(([l,v,c,pct])=>(
                    <div key={l} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"9px 0", borderBottom:`1px solid ${T.border}44` }}>
                      <div>
                        <span style={{ color:T.textDim, fontSize:13 }}>{l}</span>
                        <span style={{ marginLeft:8, fontSize:10, color:c, fontWeight:700, background:c+"22", padding:"2px 7px", borderRadius:10 }}>{Math.round(pct*100)}%</span>
                      </div>
                      <span style={{ color:c, fontWeight:700 }}>{fmtCOP(v)}</span>
                    </div>
                  ))}
                  {primaVal > 0 && (
                    <div style={{ marginTop:10, background:T.accent+"11", borderRadius:10, padding:10, fontSize:12, color:T.accent }}>
                      🎁 Prima semestral {fmtCOP(primaVal)} → abonás {fmtCOP(primaVal*strat.primaPct)} a {priorityDebt?.name||"la deuda prioritaria"}
                    </div>
                  )}
                </Card>

                {plan.length > 0 && (
                  <Card>
                    <div style={{ fontWeight:700, marginBottom:14, color:T.mint }}>
                      📅 Proyección mes a mes
                      {debtFreeLabel && <span style={{ color:T.accent, marginLeft:10, fontSize:12 }}>· Libre en {debtFreeLabel}</span>}
                    </div>
                    <div style={{ maxHeight:300, overflowY:"auto" }}>
                      {plan.map((p,i)=>{
                        const done = p.total <= 100;
                        return (
                          <div key={i} style={{ display:"flex", alignItems:"center", gap:12, padding:"8px 0", borderBottom:`1px solid ${T.border}33` }}>
                            <div style={{ width:8, height:8, borderRadius:"50%", flexShrink:0, background:done?T.mint:(p.total<totalDebt*0.5?T.accent:T.red) }} />
                            <div style={{ flex:1, fontSize:12, color:T.textDim }}>{p.label}</div>
                            <div style={{ fontWeight:700, fontSize:13, color:done?T.mint:T.text }}>{done?"🎉 ¡Libre!":fmtCOP(p.total)}</div>
                          </div>
                        );
                      })}
                    </div>
                  </Card>
                )}

                {surplus <= 0 && (
                  <Card style={{ border:`1px solid ${T.red}44`, background:T.red+"08", marginTop:12 }}>
                    <div style={{ color:T.red, fontWeight:700, marginBottom:6 }}>⚠️ Sobrante insuficiente</div>
                    <div style={{ color:T.textDim, fontSize:13, lineHeight:1.6 }}>Con los datos ingresados el sobrante es {fmtCOP(surplus)}. Revisa tus gastos fijos o cuotas mínimas en Perfil.</div>
                  </Card>
                )}
              </>
            )}
          </div>
        )}

        {/* ── GASTOS ── */}
        {tab === "gastos" && (
          <div>
            <div style={{ fontWeight:800, fontSize:18, marginBottom:4 }}>Registro de gastos</div>
            <div style={{ color:T.muted, fontSize:13, marginBottom:20 }}>Anotá cada gasto para tener control real de tu dinero.</div>

            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:16 }}>
              <Card>
                <div style={{ fontSize:10, color:T.muted, marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Este mes</div>
                <div style={{ fontWeight:800, fontSize:22, color:thisMonthTotal>incomeN*0.3?T.red:T.mint }}>{fmtCOP(thisMonthTotal)}</div>
              </Card>
              <Card>
                <div style={{ fontSize:10, color:T.muted, marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>Registrados</div>
                <div style={{ fontWeight:800, fontSize:22, color:T.accentLight }}>{expenses.length}</div>
              </Card>
            </div>

            <Card style={{ marginBottom:16 }}>
              <div style={{ fontWeight:700, marginBottom:12, color:T.accent }}>+ Nuevo gasto</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                <TextInput label="Fecha" value={newExp.date} onChange={v=>setNewExp(e=>({...e,date:v}))} type="date" />
                <MoneyInput label="Monto" value={newExp.amount} onChange={v=>setNewExp(e=>({...e,amount:v}))} />
              </div>
              <TextInput label="Descripción" value={newExp.desc} onChange={v=>setNewExp(e=>({...e,desc:v}))} />
              <Select label="Categoría" value={newExp.cat} onChange={v=>setNewExp(e=>({...e,cat:v}))} options={CATS} />
              <Btn onClick={addExpense} disabled={!newExp.amount||!newExp.desc}>Registrar gasto</Btn>
            </Card>

            {Object.keys(expByCat).length > 0 && (
              <Card style={{ marginBottom:16 }}>
                <div style={{ fontWeight:700, marginBottom:12, color:T.accent }}>Por categoría</div>
                {Object.entries(expByCat).sort((a,b)=>b[1]-a[1]).map(([cat,amt])=>(
                  <div key={cat} style={{ marginBottom:10 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4, fontSize:13 }}>
                      <span>{cat}</span><span style={{ fontWeight:700 }}>{fmtCOP(amt)}</span>
                    </div>
                    <ProgressBar pct={thisMonthTotal>0?(amt/thisMonthTotal)*100:0} color={T.accent} />
                  </div>
                ))}
              </Card>
            )}

            {expenses.length > 0 && (
              <Card>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                  <div style={{ fontWeight:700, color:T.accent }}>Historial ({expenses.length})</div>
                  <Btn variant="ghost" onClick={exportXLSX} style={{ fontSize:11, padding:"6px 12px" }}>⬇ Excel</Btn>
                </div>
                {[...expenses].reverse().map(e=>(
                  <div key={e.id} style={{ display:"flex", alignItems:"center", gap:10, padding:"9px 0", borderBottom:`1px solid ${T.border}33` }}>
                    <div style={{ flex:1 }}>
                      <div style={{ fontSize:13, fontWeight:600 }}>{e.desc}</div>
                      <div style={{ fontSize:11, color:T.muted }}>{e.cat} · {e.date}</div>
                    </div>
                    <div style={{ fontWeight:700, color:T.red }}>-{fmtCOP(Number(e.amount)||0)}</div>
                    <button onClick={()=>removeExpense(e.id)} style={{ background:"none", border:"none", color:T.muted, cursor:"pointer", fontSize:16 }}>✕</button>
                  </div>
                ))}
              </Card>
            )}

            {expenses.length === 0 && (
              <div style={{ textAlign:"center", color:T.muted, padding:40 }}>
                <div style={{ fontSize:32, marginBottom:10 }}>📝</div>
                <div>Sin gastos registrados aún.</div>
                <div style={{ fontSize:12, marginTop:6 }}>Empieza anotando lo que gastás hoy.</div>
              </div>
            )}
          </div>
        )}

        {/* ── IA ── */}
        {tab === "ia" && (
          <div>
            <div style={{ fontWeight:800, fontSize:18, marginBottom:4 }}>Consejero IA</div>
            <div style={{ color:T.muted, fontSize:13, marginBottom:16 }}>Preguntame lo que quieras. Respondo según tu situación real.</div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:16 }}>
              {["¿Cómo salgo más rápido de mis deudas?","¿Cuánto debería ahorrar este mes?","Dame tips de ingreso extra","¿Qué estrategia me conviene más?"].map(q=>(
                <button key={q} onClick={()=>setIaQ(q)} style={{ background:T.card, border:`1px solid ${T.border}`, borderRadius:20, padding:"6px 12px", color:T.textDim, fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>{q}</button>
              ))}
            </div>
            <Card style={{ marginBottom:12, minHeight:300, maxHeight:400, overflowY:"auto", padding:16 }}>
              {iaHistory.length === 0 && (
                <div style={{ color:T.muted, textAlign:"center", paddingTop:50 }}>
                  <LogoSVG size={56} />
                  <div style={{ marginTop:12, fontWeight:700, color:T.accent }}>Centro Conciencia Financiera</div>
                  <div style={{ fontSize:12, marginTop:6 }}>Conozco tu situación completa. Pregúntame lo que necesités.</div>
                </div>
              )}
              {iaHistory.map((m,i)=>(
                <div key={i} style={{ marginBottom:14, display:"flex", justifyContent:m.role==="user"?"flex-end":"flex-start" }}>
                  <div style={{ maxWidth:"85%", padding:"10px 14px", borderRadius:m.role==="user"?"16px 16px 4px 16px":"16px 16px 16px 4px", background:m.role==="user"?T.accent+"22":T.surface, border:`1px solid ${m.role==="user"?T.accent+"44":T.border}`, fontSize:13, lineHeight:1.6, color:T.text, whiteSpace:"pre-wrap" }}>
                    {m.role==="ai" && <span style={{ fontSize:10, color:T.accent, display:"block", marginBottom:4, letterSpacing:1, textTransform:"uppercase" }}>● Centro IA</span>}
                    {m.text}
                  </div>
                </div>
              ))}
              {iaLoading && (
                <div style={{ display:"flex", gap:6, padding:10 }}>
                  {[0,1,2].map(i=><div key={i} style={{ width:8, height:8, borderRadius:"50%", background:T.accent, animation:`bounce 1s ${i*0.2}s infinite` }} />)}
                </div>
              )}
              <div ref={chatEndRef} />
            </Card>
            <div style={{ display:"flex", gap:10 }}>
              <input value={iaQ} onChange={e=>setIaQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&sendIA()} placeholder="Escribe tu pregunta..."
                style={{ flex:1, background:T.card, border:`1px solid ${T.border}`, borderRadius:10, padding:"12px 14px", color:T.text, fontFamily:"inherit", fontSize:14, outline:"none" }} />
              <Btn onClick={sendIA} disabled={!iaQ.trim()||iaLoading}>Enviar</Btn>
            </div>
          </div>
        )}

        {/* ── REPORTE ── */}
        {tab === "reporte" && (
          <div>
            <div style={{ fontWeight:800, fontSize:18, marginBottom:4 }}>Reporte financiero</div>
            <div style={{ color:T.muted, fontSize:13, marginBottom:20 }}>Vista completa de tu situación y progreso.</div>

            {/* KPIs */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:16 }}>
              {[
                {label:"Ingreso mensual", value:fmtCOP(incomeN), color:T.mint},
                {label:"Deuda total",     value:fmtCOP(totalDebt), color:T.red},
                {label:"Sobrante",        value:fmtCOP(surplus), color:surplus>=0?T.mint:T.red},
                {label:"Gastos este mes", value:fmtCOP(thisMonthTotal), color:T.accentLight},
              ].map(({label,value,color})=>(
                <Card key={label}>
                  <div style={{ fontSize:10, color:T.muted, marginBottom:4, textTransform:"uppercase", letterSpacing:1 }}>{label}</div>
                  <div style={{ fontWeight:800, fontSize:18, color }}>{value}</div>
                </Card>
              ))}
            </div>

            {/* Distribución sobrante */}
            {incomeN > 0 && (
              <Card style={{ marginBottom:16 }}>
                <div style={{ fontWeight:700, marginBottom:12, color:T.accent }}>💡 Distribución mensual recomendada</div>
                {[
                  [`Abono extra → ${priorityDebt?.name||"deuda prioritaria"}`, debtExtra, T.red, strat.debtPct],
                  ["Fondo de emergencia", emergencyAmt, T.mint, strat.emergencyPct],
                  ["Colchón personal", personalAmt, T.muted, strat.personalPct],
                ].map(([l,v,c,pct])=>(
                  <div key={l} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"8px 0", borderBottom:`1px solid ${T.border}44` }}>
                    <div>
                      <span style={{ color:T.textDim, fontSize:13 }}>{l}</span>
                      <span style={{ marginLeft:8, fontSize:10, color:c, fontWeight:700 }}>{Math.round(pct*100)}%</span>
                    </div>
                    <span style={{ color:c, fontWeight:700 }}>{fmtCOP(v)}</span>
                  </div>
                ))}
              </Card>
            )}

            {/* Deudas */}
            {debts.some(d=>Number(d.balance)>0) && (
              <Card style={{ marginBottom:16 }}>
                <div style={{ fontWeight:700, marginBottom:12, color:T.accent }}>💳 Tus deudas (orden de ataque)</div>
                {sortedDebts.map((d,i)=>(
                  <div key={i} style={{ padding:"10px 0", borderBottom:`1px solid ${T.border}33` }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                      <div>
                        <span style={{ fontWeight:600 }}>{d.name || `Deuda ${i+1}`}</span>
                        {i===0 && <span style={{ marginLeft:8, fontSize:10, color:T.red, fontWeight:700 }}>🔥 Prioridad</span>}
                      </div>
                      <span style={{ color:T.red, fontWeight:700 }}>{fmtCOP(Number(d.balance)||0)}</span>
                    </div>
                    <div style={{ display:"flex", gap:16, fontSize:11, color:T.muted }}>
                      <span>Tasa: {d.rate||0}% E.A.</span>
                      <span>Mín: {fmtCOP(Number(d.minPayment)||0)}</span>
                      {i===0 && debtExtra>0 && <span style={{ color:T.red }}>+{fmtCOP(debtExtra)} extra</span>}
                    </div>
                  </div>
                ))}
              </Card>
            )}

            {/* Gastos por mes */}
            {Object.keys(expByMonth).length > 0 && (
              <Card style={{ marginBottom:16 }}>
                <div style={{ fontWeight:700, marginBottom:12, color:T.accent }}>📅 Gastos por mes</div>
                {Object.entries(expByMonth).sort().reverse().map(([month,total])=>(
                  <div key={month} style={{ display:"flex", justifyContent:"space-between", padding:"8px 0", borderBottom:`1px solid ${T.border}33`, fontSize:13 }}>
                    <span style={{ color:T.textDim }}>{month}</span>
                    <span style={{ fontWeight:700 }}>{fmtCOP(total)}</span>
                  </div>
                ))}
              </Card>
            )}

            {/* Proyección libre */}
            {debtFreeLabel && (
              <Card style={{ border:`2px solid ${T.mint}44`, background:T.mint+"08", marginBottom:16, textAlign:"center", padding:24 }}>
                <div style={{ fontSize:32, marginBottom:8 }}>🏆</div>
                <div style={{ fontWeight:800, fontSize:20, color:T.mint, marginBottom:4 }}>{debtFreeLabel}</div>
                <div style={{ fontSize:13, color:T.textDim }}>Fecha proyectada de libertad financiera total</div>
              </Card>
            )}

            {/* Reglas dinámicas */}
            <Card style={{ border:`1px solid ${T.mint}33`, marginBottom:16 }}>
              <div style={{ fontWeight:700, marginBottom:4, color:T.mint }}>✅ Tus reglas personalizadas</div>
              <div style={{ fontSize:10, color:T.muted, marginBottom:12, letterSpacing:0.5 }}>Estrategia {strat.label} · basadas en tus datos reales</div>
              {dynamicRules().map((rule,i)=>(
                <div key={i} style={{ display:"flex", gap:10, padding:"8px 0", borderBottom:`1px solid ${T.border}33`, fontSize:13 }}>
                  <span style={{ color:T.accent, flexShrink:0 }}>→</span>
                  <span style={{ color:T.textDim, lineHeight:1.5 }}>{rule}</span>
                </div>
              ))}
            </Card>

            {/* Export */}
            <Btn onClick={exportXLSX} style={{ width:"100%" }}>⬇ Exportar reporte completo Excel</Btn>
          </div>
        )}

      </div>

      {/* FOOTER */}
      <div style={{ textAlign:"center", padding:"24px 16px", borderTop:`1px solid ${T.border}`, marginTop:20 }}>
        <LogoSVG size={28} />
        <div style={{ fontSize:9, color:T.muted, letterSpacing:2, textTransform:"uppercase", marginTop:6 }}>Centro · Conciencia Financiera</div>
      </div>

      <style>{`
        @keyframes bounce{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}
        input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}
        ::-webkit-scrollbar{width:4px}
        ::-webkit-scrollbar-track{background:transparent}
        ::-webkit-scrollbar-thumb{background:${T.border};border-radius:4px}
        select option{background:${T.surface}}
      `}</style>
    </div>
  );
}
