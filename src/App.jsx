import { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, AreaChart, Area, CartesianGrid
} from "recharts";
import {
  Users, TrendingUp, TrendingDown, DollarSign, BarChart2,
  Upload, Search, ChevronUp, ChevronDown, LayoutDashboard,
  CreditCard, Table2, Menu, X, Award, Briefcase, Activity,
  RefreshCw, Download, Moon, Sun, Shield, ArrowUpRight,
  ArrowDownRight, Minus, Filter, ChevronLeft, ChevronRight,
  Building2, UserCheck, UserX, Zap, Star, Clock
} from "lucide-react";

const C = {
  emerald: "#10B981",
  cyan: "#06B6D4",
  purple: "#8B5CF6",
  amber: "#F59E0B",
  rose: "#F43F5E",
  indigo: "#6366F1",
  bg: "#080C14",
  surface: "#0D1117",
  card: "#111827",
  border: "#1F2937",
  border2: "#374151",
  text: "#F9FAFB",
  textMid: "#9CA3AF",
  textMuted: "#4B5563",
};

const SERIES = [C.emerald, C.cyan, C.purple, C.amber, C.rose, C.indigo];

const SAMPLE = [
  { "Employee ID": "EMP001", "Employee Name": "Aarav Sharma", Gender: "Female", Department: "Operations", Designation: "Senior Executive", "Join Date": "2025-10-21", Salary: 701459, Status: "Active" },
  { "Employee ID": "EMP002", "Employee Name": "Vivaan Patel", Gender: "Male", Department: "Sales", Designation: "Manager", "Join Date": "2025-09-22", Salary: 936806, Status: "Active" },
  { "Employee ID": "EMP003", "Employee Name": "Aditya Rao", Gender: "Female", Department: "Sales", Designation: "Analyst", "Join Date": "2023-10-07", Salary: 766843, Status: "Active" },
  { "Employee ID": "EMP004", "Employee Name": "Arjun Kumar", Gender: "Female", Department: "Sales", Designation: "Senior Executive", "Join Date": "2025-03-10", Salary: 1268486, Status: "Active" },
  { "Employee ID": "EMP005", "Employee Name": "Karthik Reddy", Gender: "Female", Department: "HR", Designation: "Analyst", "Join Date": "2025-07-11", Salary: 414887, Status: "Active" },
  { "Employee ID": "EMP006", "Employee Name": "Rohan Mehta", Gender: "Female", Department: "HR", Designation: "Analyst", "Join Date": "2025-08-27", Salary: 688433, Status: "Active" },
  { "Employee ID": "EMP007", "Employee Name": "Priya Nair", Gender: "Female", Department: "Finance", Designation: "Manager", "Join Date": "2022-05-14", Salary: 1050000, Status: "Active" },
  { "Employee ID": "EMP008", "Employee Name": "Rahul Gupta", Gender: "Male", Department: "Operations", Designation: "Analyst", "Join Date": "2024-01-18", Salary: 520000, Status: "Inactive" },
  { "Employee ID": "EMP009", "Employee Name": "Sneha Iyer", Gender: "Female", Department: "Finance", Designation: "Senior Executive", "Join Date": "2021-11-03", Salary: 890000, Status: "Active" },
  { "Employee ID": "EMP010", "Employee Name": "Vikram Singh", Gender: "Male", Department: "HR", Designation: "Manager", "Join Date": "2023-06-20", Salary: 750000, Status: "Active" },
  { "Employee ID": "EMP011", "Employee Name": "Meera Pillai", Gender: "Female", Department: "Finance", Designation: "Analyst", "Join Date": "2024-08-05", Salary: 610000, Status: "Active" },
  { "Employee ID": "EMP012", "Employee Name": "Dev Kapoor", Gender: "Male", Department: "Operations", Designation: "Manager", "Join Date": "2020-03-17", Salary: 980000, Status: "Active" },
  { "Employee ID": "EMP013", "Employee Name": "Ananya Bose", Gender: "Female", Department: "Sales", Designation: "Analyst", "Join Date": "2022-11-29", Salary: 555000, Status: "Inactive" },
  { "Employee ID": "EMP014", "Employee Name": "Rajan Menon", Gender: "Male", Department: "Finance", Designation: "Senior Executive", "Join Date": "2021-07-08", Salary: 820000, Status: "Active" },
  { "Employee ID": "EMP015", "Employee Name": "Kavitha Das", Gender: "Female", Department: "Operations", Designation: "Analyst", "Join Date": "2023-04-22", Salary: 480000, Status: "Active" },
];

const REQ_COLS = ["Employee ID", "Employee Name", "Gender", "Department", "Designation", "Join Date", "Salary", "Status"];

const fmt = (n) => new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(n);
const fmtCur = (n) => "₹" + fmt(n);
const fmtK = (n) => (n >= 1e7 ? "₹" + (n / 1e7).toFixed(1) + "Cr" : n >= 1e5 ? "₹" + (n / 1e5).toFixed(1) + "L" : "₹" + fmt(n));
const initials = (s) => s.split(" ").map((w) => w[0]).join("").slice(0, 2).toUpperCase();
const pct = (a, b) => (b ? (((a - b) / b) * 100).toFixed(1) : "0.0");
const AVATAR_COLORS = [C.emerald, C.cyan, C.purple, C.amber, C.rose, C.indigo];
const avatarColor = (name) => AVATAR_COLORS[name.charCodeAt(0) % AVATAR_COLORS.length];

const Spark = ({ data, color, dataKey = "v" }) => (
  <ResponsiveContainer width="100%" height={44}>
    <AreaChart data={data} margin={{ top: 4, right: 0, left: 0, bottom: 0 }}>
      <defs>
        <linearGradient id={`sg-${color.replace("#", "")}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={color} stopOpacity={0.25} />
          <stop offset="100%" stopColor={color} stopOpacity={0} />
        </linearGradient>
      </defs>
      <Area type="monotone" dataKey={dataKey} stroke={color} strokeWidth={1.5} fill={`url(#sg-${color.replace("#", "")})`} dot={false} isAnimationActive={false} />
    </AreaChart>
  </ResponsiveContainer>
);

const DarkTip = ({ active, payload, label, currency }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#1a2233", border: `1px solid ${C.border2}`, borderRadius: 8, padding: "10px 14px", fontSize: 12, boxShadow: "0 8px 32px rgba(0,0,0,0.5)" }}>
      {label && <div style={{ color: C.textMid, marginBottom: 5, fontWeight: 600 }}>{label}</div>}
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || C.emerald, display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ width: 6, height: 6, borderRadius: "50%", background: p.color || C.emerald, flexShrink: 0 }} />
          <span style={{ color: C.textMid }}>{p.name}:</span>
          <span style={{ fontWeight: 700, color: C.text }}>{currency ? fmtK(p.value) : fmt(p.value)}</span>
        </div>
      ))}
    </div>
  );
};

const KPI = ({ icon: Icon, label, value, sub, color, spark, trend, trendLabel }) => {
  const trendPos = parseFloat(trend) > 0;
  const trendNeg = parseFloat(trend) < 0;
  return (
    <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "20px 20px 14px", display: "flex", flexDirection: "column", gap: 0, position: "relative", overflow: "hidden", transition: "border-color 0.2s" }}>
      <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 2, background: color, opacity: 0.7, borderRadius: "14px 14px 0 0" }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
        <span style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase" }}>{label}</span>
        <div style={{ width: 32, height: 32, borderRadius: 8, background: `${color}18`, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icon size={15} color={color} />
        </div>
      </div>
      <div style={{ fontSize: 26, fontWeight: 700, color: C.text, lineHeight: 1, letterSpacing: "-0.02em", marginBottom: 6 }}>{value}</div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontSize: 11, color: C.textMuted }}>{sub}</span>
        {trend !== undefined && (
          <div style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 11, fontWeight: 600, color: trendPos ? C.emerald : trendNeg ? C.rose : C.textMid }}>
            {trendPos ? <ArrowUpRight size={12} /> : trendNeg ? <ArrowDownRight size={12} /> : <Minus size={12} />}
            {Math.abs(parseFloat(trend))}% {trendLabel || ""}
          </div>
        )}
      </div>
      {spark && <div style={{ marginTop: 8 }}><Spark data={spark} color={color} /></div>}
    </div>
  );
};

const Card = ({ title, subtitle, action, children, style }) => (
  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, overflow: "hidden", ...style }}>
    {(title || action) && (
      <div style={{ padding: "16px 20px 14px", display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: `1px solid ${C.border}` }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.textMid, letterSpacing: "0.04em", textTransform: "uppercase" }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>{subtitle}</div>}
        </div>
        {action}
      </div>
    )}
    <div style={{ padding: "16px 20px" }}>{children}</div>
  </div>
);

const Badge = ({ label, color, bg }) => (
  <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 20, background: bg || `${color}18`, color: color || C.textMid }}>
    <span style={{ width: 5, height: 5, borderRadius: "50%", background: color || C.textMid }} />
    {label}
  </span>
);

const Uploader = ({ onData, count }) => {
  const [drag, setDrag] = useState(false);
  const [msg, setMsg] = useState(null);
  const ref = useRef();

  const process = useCallback((file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    const r = new FileReader();
    r.onload = (e) => {
      try {
        const wb = XLSX.read(ext === "csv" ? e.target.result : new Uint8Array(e.target.result), { type: ext === "csv" ? "binary" : "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const miss = REQ_COLS.filter((c) => !Object.keys(rows[0] || {}).includes(c));
        if (miss.length) {
          setMsg({ err: true, t: `Missing: ${miss.join(", ")}` });
          return;
        }
        const clean = rows.map((r2) => ({
          ...r2,
          Salary: parseFloat(String(r2.Salary).replace(/[^0-9.]/g, "")) || 0,
          "Join Date": String(r2["Join Date"] || "").slice(0, 10),
          "Employee Name": String(r2["Employee Name"] || "").trim(),
          Department: String(r2.Department || "").trim(),
          Gender: String(r2.Gender || "").trim(),
          Status: String(r2.Status || "").trim(),
          Designation: String(r2.Designation || "").trim(),
        }));
        onData(clean);
        setMsg({ err: false, t: `${clean.length} records loaded from ${file.name}` });
      } catch (e2) {
        setMsg({ err: true, t: "Parse error: " + e2.message });
      }
    };
    ext === "csv" ? r.readAsBinaryString(file) : r.readAsArrayBuffer(file);
  }, [onData]);

  return (
    <div>
      <div onClick={() => ref.current?.click()} onDrop={(e) => { e.preventDefault(); setDrag(false); process(e.dataTransfer.files[0]); }} onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", borderRadius: 10, border: `1.5px dashed ${drag ? C.cyan : C.border2}`, cursor: "pointer", background: drag ? `${C.cyan}08` : "transparent", transition: "all 0.18s" }}>
        <div style={{ width: 34, height: 34, borderRadius: 8, background: `${C.cyan}18`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Upload size={15} color={C.cyan} />
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#CBD5E1" }}>Upload Excel or CSV</div>
          <div style={{ fontSize: 11, color: C.textMuted }}>Drag &amp; drop · Processed locally · No server upload</div>
        </div>
        <div style={{ fontSize: 11, color: C.textMuted, textAlign: "right", flexShrink: 0 }}>
          <div style={{ fontWeight: 600, color: C.emerald }}>{count}</div>
          <div>records</div>
        </div>
        <input ref={ref} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => process(e.target.files[0])} />
      </div>
      {msg && <div style={{ marginTop: 6, fontSize: 11, padding: "5px 12px", borderRadius: 6, background: msg.err ? `${C.rose}12` : `${C.emerald}10`, color: msg.err ? C.rose : C.emerald, border: `1px solid ${msg.err ? C.rose : C.emerald}28` }}>{msg.t}</div>}
    </div>
  );
};

const exportCSV = (data) => {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Employees");
  XLSX.writeFile(wb, `talentpulse_export_${Date.now()}.xlsx`);
};

const Overview = ({ data }) => {
  const total = data.length;
  const active = data.filter((r) => r.Status === "Active").length;
  const inactive = total - active;
  const payroll = data.reduce((s, r) => s + r.Salary, 0);
  const avg = total ? Math.round(payroll / total) : 0;

  const hiringByMonth = useMemo(() => {
    const m = {};
    data.forEach((r) => {
      const y = r["Join Date"]?.slice(0, 7);
      if (y) m[y] = (m[y] || 0) + 1;
    });
    return Object.entries(m).sort(([a], [b]) => a.localeCompare(b)).map(([k, v]) => ({ k, v }));
  }, [data]);

  const spark = hiringByMonth.slice(-8).map((x) => ({ v: x.v }));

  const byDept = useMemo(() => {
    const m = {};
    data.forEach((r) => {
      m[r.Department] = (m[r.Department] || 0) + 1;
    });
    return Object.entries(m).map(([n, c]) => ({ n, c })).sort((a, b) => b.c - a.c);
  }, [data]);

  const byGender = useMemo(() => {
    const m = {};
    data.forEach((r) => {
      m[r.Gender] = (m[r.Gender] || 0) + 1;
    });
    return Object.entries(m).map(([n, v]) => ({ n, v }));
  }, [data]);

  const byStatus = [
    { n: "Active", v: active, color: C.emerald },
    { n: "Inactive", v: inactive, color: C.rose },
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 }}>
        <KPI icon={Users} label="Total workforce" value={total} sub={`${active} active · ${inactive} inactive`} color={C.cyan} spark={spark} />
        <KPI icon={UserCheck} label="Active rate" value={`${Math.round((active / total) * 100 || 0)}%`} sub="of total headcount" color={C.emerald} trend={pct(active, Math.round(total * 0.85))} trendLabel="vs target" />
        <KPI icon={DollarSign} label="Annual payroll" value={fmtK(payroll)} sub="total compensation" color={C.amber} spark={byDept.map((d) => ({ v: d.c * 50000 }))} />
        <KPI icon={TrendingUp} label="Avg salary" value={fmtK(avg)} sub="company-wide" color={C.purple} trend={pct(avg, 600000)} trendLabel="vs ₹6L" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 14 }}>
        <Card title="Headcount by department" subtitle="All employees including inactive">
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={byDept} margin={{ top: 4, right: 4, left: -22, bottom: 0 }}>
              <CartesianGrid stroke={C.border} vertical={false} strokeDasharray="3 3" />
              <XAxis dataKey="n" tick={{ fill: C.textMuted, fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: C.textMuted, fontSize: 11 }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip content={<DarkTip />} />
              <Bar dataKey="c" name="Employees" radius={[5, 5, 0, 0]} maxBarSize={44}>
                {byDept.map((_, i) => <Cell key={i} fill={SERIES[i % SERIES.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card title="Gender breakdown" subtitle="Diversity ratio">
          <ResponsiveContainer width="100%" height={230}>
            <PieChart>
              <Pie data={byGender} cx="50%" cy="46%" innerRadius={62} outerRadius={92} dataKey="v" nameKey="n" paddingAngle={4} strokeWidth={0}>
                {byGender.map((_, i) => <Cell key={i} fill={SERIES[i % SERIES.length]} />)}
              </Pie>
              <Tooltip content={<DarkTip />} />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ display: "flex", justifyContent: "center", gap: 16, marginTop: -4 }}>
            {byGender.map((g, i) => (
              <div key={g.n} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11 }}>
                <span style={{ width: 8, height: 8, borderRadius: 2, background: SERIES[i % SERIES.length] }} />
                <span style={{ color: C.textMid }}>{g.n}</span>
                <span style={{ color: C.text, fontWeight: 600 }}>{g.v}</span>
              </div>
            ))}
          </div>
        </Card>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <Card title="Hiring trend" subtitle="Monthly hires over time">
          <ResponsiveContainer width="100%" height={190}>
            <AreaChart data={hiringByMonth} margin={{ top: 4, right: 4, left: -22, bottom: 0 }}>
              <defs>
                <linearGradient id="htg" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={C.cyan} stopOpacity={0.3} />
                  <stop offset="100%" stopColor={C.cyan} stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={C.border} vertical={false} strokeDasharray="3 3" />
              <XAxis dataKey="k" tick={{ fill: C.textMuted, fontSize: 10 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: C.textMuted, fontSize: 10 }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip content={<DarkTip />} />
              <Area type="monotone" dataKey="v" name="Hires" stroke={C.cyan} strokeWidth={2} fill="url(#htg)" dot={{ r: 3, fill: C.cyan, strokeWidth: 0 }} />
            </AreaChart>
          </ResponsiveContainer>
        </Card>

        <Card title="Active vs inactive" subtitle="Current workforce status">
          <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 4 }}>
            {byStatus.map((s) => {
              const pct2 = total ? Math.round((s.v / total) * 100) : 0;
              return (
                <div key={s.n}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                    <span style={{ fontSize: 12, color: C.textMid }}>{s.n}</span>
                    <span style={{ fontSize: 12, fontWeight: 700, color: s.color }}>{s.v} <span style={{ color: C.textMuted, fontWeight: 400 }}>({pct2}%)</span></span>
                  </div>
                  <div style={{ height: 8, borderRadius: 4, background: C.border, overflow: "hidden" }}>
                    <div style={{ height: "100%", width: `${pct2}%`, background: s.color, borderRadius: 4, transition: "width 0.6s ease" }} />
                  </div>
                </div>
              );
            })}
            <div style={{ marginTop: 8, padding: "12px 14px", borderRadius: 10, background: `${C.emerald}08`, border: `1px solid ${C.emerald}22` }}>
              <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 4 }}>Retention health</div>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.emerald }}>{Math.round((active / total) * 100 || 0)}%</div>
              <div style={{ fontSize: 11, color: C.textMuted }}>active workforce ratio</div>
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
};

const Compensation = ({ data }) => {
  const avgByDept = useMemo(() => {
    const m = {}, cnt = {};
    data.forEach((r) => {
      m[r.Department] = (m[r.Department] || 0) + r.Salary;
      cnt[r.Department] = (cnt[r.Department] || 0) + 1;
    });
    return Object.entries(m).map(([n, t]) => ({ n, avg: Math.round(t / cnt[n]) })).sort((a, b) => b.avg - a.avg);
  }, [data]);

  const byDesig = useMemo(() => {
    const m = {}, cnt = {};
    data.forEach((r) => {
      m[r.Designation] = (m[r.Designation] || 0) + r.Salary;
      cnt[r.Designation] = (cnt[r.Designation] || 0) + 1;
    });
    return Object.entries(m).map(([n, t]) => ({ n, avg: Math.round(t / cnt[n]), cnt: cnt[n] })).sort((a, b) => b.avg - a.avg);
  }, [data]);

  const top5 = useMemo(() => [...data].sort((a, b) => b.Salary - a.Salary).slice(0, 5), [data]);
  const bottom5 = useMemo(() => [...data].sort((a, b) => a.Salary - b.Salary).slice(0, 5), [data]);

  const salRange = useMemo(() => {
    const salaries = data.map((r) => r.Salary).sort((a, b) => a - b);
    const buckets = [0, 3e5, 5e5, 7e5, 9e5, 12e5, Infinity];
    const labels = ["<3L", "3–5L", "5–7L", "7–9L", "9–12L", "12L+"]; 
    return labels.map((lbl, i) => ({ lbl, count: salaries.filter((s) => s >= buckets[i] && s < buckets[i + 1]).length }));
  }, [data]);

  const companyAvg = data.length ? Math.round(data.reduce((s, r) => s + r.Salary, 0) / data.length) : 0;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
        <KPI icon={DollarSign} label="Highest salary" value={fmtK(Math.max(...data.map((r) => r.Salary)))} sub={top5[0]?.["Employee Name"]} color={C.emerald} />
        <KPI icon={TrendingUp} label="Avg salary" value={fmtK(companyAvg)} sub="company median" color={C.cyan} />
        <KPI icon={TrendingDown} label="Lowest salary" value={fmtK(Math.min(...data.map((r) => r.Salary)))} sub={bottom5[0]?.["Employee Name"]} color={C.amber} />
        <KPI icon={Zap} label="Pay spread" value={fmtK(Math.max(...data.map((r) => r.Salary)) - Math.min(...data.map((r) => r.Salary)))} sub="max minus min" color={C.purple} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <Card title="Avg salary by department" subtitle="Sorted highest to lowest">
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={avgByDept} layout="vertical" margin={{ top: 0, right: 12, left: 56, bottom: 0 }}>
              <CartesianGrid stroke={C.border} horizontal={false} strokeDasharray="3 3" />
              <XAxis type="number" tick={{ fill: C.textMuted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={(v) => "₹" + (v / 1e5).toFixed(0) + "L"} />
              <YAxis type="category" dataKey="n" tick={{ fill: C.textMid, fontSize: 11 }} axisLine={false} tickLine={false} width={52} />
              <Tooltip content={<DarkTip currency />} />
              <Bar dataKey="avg" name="Avg salary" radius={[0, 5, 5, 0]} maxBarSize={20}>
                {avgByDept.map((_, i) => <Cell key={i} fill={SERIES[i % SERIES.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card title="Salary distribution" subtitle="Employees per salary band">
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={salRange} margin={{ top: 4, right: 4, left: -22, bottom: 0 }}>
              <CartesianGrid stroke={C.border} vertical={false} strokeDasharray="3 3" />
              <XAxis dataKey="lbl" tick={{ fill: C.textMuted, fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: C.textMuted, fontSize: 11 }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip content={<DarkTip />} />
              <Bar dataKey="count" name="Employees" radius={[5, 5, 0, 0]} fill={C.purple} maxBarSize={42} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <Card title="Compensation by designation" subtitle="Average pay per role, size = headcount">
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {byDesig.map((d, i) => {
            const barW = Math.round((d.avg / Math.max(...byDesig.map((x) => x.avg))) * 100);
            return (
              <div key={d.n} style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{ width: 130, fontSize: 12, color: C.textMid, textAlign: "right", flexShrink: 0 }}>{d.n}</div>
                <div style={{ flex: 1, height: 10, borderRadius: 5, background: C.border, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${barW}%`, background: SERIES[i % SERIES.length], borderRadius: 5, transition: "width 0.5s ease" }} />
                </div>
                <div style={{ width: 80, fontSize: 12, fontWeight: 600, color: SERIES[i % SERIES.length], flexShrink: 0 }}>{fmtK(d.avg)}</div>
                <div style={{ width: 36, fontSize: 11, color: C.textMuted, textAlign: "right", flexShrink: 0 }}>{d.cnt} ppl</div>
              </div>
            );
          })}
        </div>
      </Card>

      <Card title="Top 5 earners" subtitle="Highest compensated employees">
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {top5.map((emp, i) => {
            const aColor = avatarColor(emp["Employee Name"]);
            return (
              <div key={emp["Employee ID"]} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 14px", borderRadius: 10, background: `${C.bg}`, border: `1px solid ${C.border}`, transition: "border-color 0.15s" }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: SERIES[i], width: 20, textAlign: "center", flexShrink: 0 }}>#{i + 1}</div>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: `${aColor}22`, border: `1.5px solid ${aColor}44`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 700, color: aColor, flexShrink: 0 }}>
                  {initials(emp["Employee Name"])}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.text, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{emp["Employee Name"]}</div>
                  <div style={{ fontSize: 11, color: C.textMuted }}>{emp.Designation} · {emp.Department}</div>
                </div>
                <div style={{ textAlign: "right", flexShrink: 0 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, color: C.emerald }}>{fmtCur(emp.Salary)}</div>
                  <div style={{ fontSize: 10, color: C.textMuted }}>{emp.Status === "Active" ? "Active" : "Inactive"}</div>
                </div>
              </div>
            );
          })}
        </div>
      </Card>
    </div>
  );
};

const Directory = ({ data }) => {
  const [q, setQ] = useState("");
  const [dept, setDept] = useState("All");
  const [gen, setGen] = useState("All");
  const [stat, setStat] = useState("All");
  const [desig, setDesig] = useState("All");
  const [sortK, setSortK] = useState("Salary");
  const [sortD, setSortD] = useState("desc");
  const [page, setPage] = useState(1);
  const PER = 10;

  const depts = useMemo(() => ["All", ...new Set(data.map((r) => r.Department))], [data]);
  const gens = useMemo(() => ["All", ...new Set(data.map((r) => r.Gender))], [data]);
  const desigs = useMemo(() => ["All", ...new Set(data.map((r) => r.Designation))], [data]);

  const filtered = useMemo(() => {
    let rows = data;
    if (q) {
      const lq = q.toLowerCase();
      rows = rows.filter((r) => r["Employee Name"].toLowerCase().includes(lq) || r["Employee ID"].toLowerCase().includes(lq) || r.Designation.toLowerCase().includes(lq));
    }
    if (dept !== "All") rows = rows.filter((r) => r.Department === dept);
    if (gen !== "All") rows = rows.filter((r) => r.Gender === gen);
    if (stat !== "All") rows = rows.filter((r) => r.Status === stat);
    if (desig !== "All") rows = rows.filter((r) => r.Designation === desig);
    return [...rows].sort((a, b) => {
      const av = sortK === "Salary" ? a.Salary : (a["Join Date"] || "");
      const bv = sortK === "Salary" ? b.Salary : (b["Join Date"] || "");
      return av < bv ? (sortD === "asc" ? -1 : 1) : av > bv ? (sortD === "asc" ? 1 : -1) : 0;
    });
  }, [data, q, dept, gen, stat, desig, sortK, sortD]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PER));
  const paged = filtered.slice((page - 1) * PER, page * PER);

  const toggleSort = (k) => {
    if (sortK === k) setSortD((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortK(k); setSortD("desc"); }
    setPage(1);
  };
  const reset = () => { setQ(""); setDept("All"); setGen("All"); setStat("All"); setDesig("All"); setPage(1); };

  const selS = { background: C.surface, border: `1px solid ${C.border}`, borderRadius: 7, padding: "6px 10px", color: C.textMid, fontSize: 12, outline: "none", cursor: "pointer" };

  const SortArrow = ({ k }) => sortK === k ? (sortD === "asc" ? <ChevronUp size={12} color={C.cyan} /> : <ChevronDown size={12} color={C.cyan} />) : <ChevronDown size={12} style={{ opacity: 0.25 }} />;

  const TH = ({ label, sortable, k, w }) => (
    <th onClick={sortable ? () => toggleSort(k) : undefined} style={{ padding: "10px 14px", textAlign: "left", fontSize: 10, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", color: sortK === k ? C.cyan : C.textMuted, background: "#0a0e18", whiteSpace: "nowrap", width: w, cursor: sortable ? "pointer" : "default", userSelect: "none", borderBottom: `1px solid ${C.border}` }}>
      <span style={{ display: "flex", alignItems: "center", gap: 4 }}>{label}{sortable && <SortArrow k={k} />}</span>
    </th>
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        {[
          { label: "Total", val: data.length, color: C.cyan },
          { label: "Filtered", val: filtered.length, color: C.purple },
          { label: "Active", val: data.filter((r) => r.Status === "Active").length, color: C.emerald },
          { label: "Inactive", val: data.filter((r) => r.Status === "Inactive").length, color: C.rose },
        ].map((x) => (
          <div key={x.label} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 14px", borderRadius: 8, background: C.card, border: `1px solid ${C.border}`, flex: "1 1 auto" }}>
            <span style={{ fontSize: 18, fontWeight: 700, color: x.color }}>{x.val}</span>
            <span style={{ fontSize: 11, color: C.textMuted }}>{x.label}</span>
          </div>
        ))}
      </div>

      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "14px 16px" }}>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
          <div style={{ position: "relative", flex: "1 1 200px" }}>
            <Search size={13} style={{ position: "absolute", left: 9, top: "50%", transform: "translateY(-50%)", color: C.textMuted }} />
            <input value={q} onChange={(e) => { setQ(e.target.value); setPage(1); }} placeholder="Search name, ID, or designation…" style={{ ...selS, paddingLeft: 28, width: "100%", boxSizing: "border-box" }} />
          </div>
          {[
            { val: dept, set: setDept, opts: depts, ph: "Department" },
            { val: gen, set: setGen, opts: gens, ph: "Gender" },
            { val: stat, set: setStat, opts: ["All", "Active", "Inactive"], ph: "Status" },
            { val: desig, set: setDesig, opts: desigs, ph: "Designation" },
          ].map((f, i) => (
            <select key={i} value={f.val} onChange={(e) => { f.set(e.target.value); setPage(1); }} style={selS}>
              {f.opts.map((o) => <option key={o}>{o}</option>)}
            </select>
          ))}
          <button onClick={reset} style={{ ...selS, display: "flex", alignItems: "center", gap: 5, color: C.textMuted, background: "transparent" }}>
            <RefreshCw size={12} /> Reset
          </button>
          <button onClick={() => exportCSV(filtered)} style={{ ...selS, display: "flex", alignItems: "center", gap: 5, color: C.cyan, background: `${C.cyan}10`, borderColor: `${C.cyan}30` }}>
            <Download size={12} /> Export
          </button>
        </div>
      </div>

      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, tableLayout: "auto" }}>
            <thead>
              <tr>
                <TH label="Employee" w={200} />
                <TH label="ID" w={90} />
                <TH label="Gender" w={80} />
                <TH label="Department" w={120} />
                <TH label="Designation" w={150} />
                <TH label="Join Date" sortable k="Join Date" w={105} />
                <TH label="Salary" sortable k="Salary" w={110} />
                <TH label="Status" w={90} />
              </tr>
            </thead>
            <tbody>
              {paged.length === 0 && (
                <tr><td colSpan={8} style={{ padding: "40px", textAlign: "center", color: C.textMuted, fontSize: 13 }}>No records match your filters.</td></tr>
              )}
              {paged.map((row, i) => {
                const aColor = avatarColor(row["Employee Name"]);
                return (
                  <tr key={row["Employee ID"]} style={{ borderBottom: `1px solid ${C.border}`, background: i % 2 === 0 ? "transparent" : "rgba(10,14,24,0.5)", transition: "background 0.12s" }}>
                    <td style={{ padding: "10px 14px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
                        <div style={{ width: 30, height: 30, borderRadius: 8, background: `${aColor}22`, border: `1.5px solid ${aColor}44`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700, color: aColor, flexShrink: 0 }}>{initials(row["Employee Name"])}</div>
                        <span style={{ fontWeight: 600, color: C.text, fontSize: 12 }}>{row["Employee Name"]}</span>
                      </div>
                    </td>
                    <td style={{ padding: "10px 14px", color: C.textMuted, fontFamily: "monospace", fontSize: 11 }}>{row["Employee ID"]}</td>
                    <td style={{ padding: "10px 14px", color: C.textMid, fontSize: 12 }}>{row.Gender}</td>
                    <td style={{ padding: "10px 14px" }}><span style={{ fontSize: 11, fontWeight: 600, padding: "3px 8px", borderRadius: 5, background: `${C.purple}18`, color: C.purple }}>{row.Department}</span></td>
                    <td style={{ padding: "10px 14px", color: C.textMid, fontSize: 12 }}>{row.Designation}</td>
                    <td style={{ padding: "10px 14px", color: C.textMuted, fontSize: 11 }}>{row["Join Date"]}</td>
                    <td style={{ padding: "10px 14px", fontWeight: 700, color: C.emerald, fontSize: 12 }}>{fmtCur(row.Salary)}</td>
                    <td style={{ padding: "10px 14px" }}><Badge label={row.Status} color={row.Status === "Active" ? C.emerald : C.rose} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 16px", borderTop: `1px solid ${C.border}`, background: "#0a0e18" }}>
          <span style={{ fontSize: 11, color: C.textMuted }}>{filtered.length} records · showing {(page - 1) * PER + 1}–{Math.min(page * PER, filtered.length)}</span>
          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
            <button onClick={() => setPage((p) => Math.max(1, p - 1))} disabled={page === 1} style={{ ...selS, padding: "4px 8px", display: "flex", background: "transparent", opacity: page === 1 ? 0.3 : 1 }}>
              <ChevronLeft size={14} />
            </button>
            {[...Array(totalPages)].map((_, i) => (
              <button key={i} onClick={() => setPage(i + 1)} style={{ ...selS, padding: "4px 10px", minWidth: 30, background: page === i + 1 ? C.cyan : "transparent", color: page === i + 1 ? "#0B0F19" : C.textMuted, fontWeight: page === i + 1 ? 700 : 400, borderColor: page === i + 1 ? C.cyan : C.border }}>
                {i + 1}
              </button>
            )).slice(Math.max(0, page - 3), page + 2)}
            <button onClick={() => setPage((p) => Math.min(totalPages, p + 1))} disabled={page === totalPages} style={{ ...selS, padding: "4px 8px", display: "flex", background: "transparent", opacity: page === totalPages ? 0.3 : 1 }}>
              <ChevronRight size={14} />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

const NAV = [
  { id: "overview", label: "Executive overview", sub: "KPIs & trends", icon: LayoutDashboard },
  { id: "compensation", label: "Compensation", sub: "Pay analytics", icon: CreditCard },
  { id: "directory", label: "Talent directory", sub: "Employee records", icon: Table2 },
];

export default function TalentPulse() {
  const [data, setData] = useState(SAMPLE);
  const [page, setPage] = useState("overview");
  const [open, setOpen] = useState(true);
  const now = new Date().toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short", year: "numeric" });

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif", fontSize: 14 }}>
      <aside style={{ width: open ? 248 : 62, flexShrink: 0, background: C.surface, borderRight: `1px solid ${C.border}`, display: "flex", flexDirection: "column", position: "sticky", top: 0, height: "100vh", overflowY: "auto", overflowX: "hidden", transition: "width 0.22s cubic-bezier(.4,0,.2,1)" }}>
        <div style={{ padding: "18px 14px 14px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: `linear-gradient(135deg,${C.emerald},${C.cyan})`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            <BarChart2 size={17} color="#060d18" />
          </div>
          {open && (
            <div style={{ overflow: "hidden", minWidth: 0 }}>
              <div style={{ fontSize: 15, fontWeight: 700, color: C.text, letterSpacing: "-0.02em", whiteSpace: "nowrap" }}>TalentPulse</div>
              <div style={{ fontSize: 10, color: C.textMuted, marginTop: 1 }}>HR Analytics Portal</div>
            </div>
          )}
        </div>

        <nav style={{ flex: 1, padding: "10px 8px", display: "flex", flexDirection: "column", gap: 2 }}>
          {open && <div style={{ fontSize: 10, fontWeight: 600, letterSpacing: "0.1em", textTransform: "uppercase", color: C.textMuted, padding: "8px 8px 4px" }}>Navigation</div>}
          {NAV.map((n) => {
            const Ic = n.icon;
            const active = page === n.id;
            return (
              <button key={n.id} onClick={() => setPage(n.id)} title={!open ? n.label : undefined} style={{ display: "flex", alignItems: "center", gap: 10, padding: open ? "9px 10px" : "10px", borderRadius: 9, border: "none", cursor: "pointer", width: "100%", textAlign: "left", background: active ? `${C.cyan}14` : "transparent", transition: "background 0.14s", justifyContent: open ? "flex-start" : "center" }}>
                <div style={{ width: 28, height: 28, borderRadius: 7, display: "flex", alignItems: "center", justifyContent: "center", background: active ? `${C.cyan}20` : "transparent", flexShrink: 0 }}>
                  <Ic size={15} color={active ? C.cyan : C.textMuted} />
                </div>
                {open && (
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: active ? 600 : 400, color: active ? C.cyan : C.textMid, whiteSpace: "nowrap" }}>{n.label}</div>
                    <div style={{ fontSize: 10, color: C.textMuted }}>{n.sub}</div>
                  </div>
                )}
                {open && active && <div style={{ marginLeft: "auto", width: 4, height: 4, borderRadius: "50%", background: C.cyan, flexShrink: 0 }} />}
              </button>
            );
          })}
        </nav>

        <div style={{ padding: "10px 8px", borderTop: `1px solid ${C.border}`, flexShrink: 0 }}>
          <button onClick={() => setOpen((v) => !v)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, cursor: "pointer", background: "transparent", color: C.textMuted, width: "100%", justifyContent: open ? "flex-start" : "center", transition: "all 0.14s" }}>
            {open ? <><ChevronLeft size={13} /><span style={{ fontSize: 11 }}>Collapse sidebar</span></> : <ChevronRight size={13} />}
          </button>
        </div>
      </aside>

      <main style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
        <header style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, background: `${C.surface}ee`, backdropFilter: "blur(16px)", position: "sticky", top: 0, zIndex: 10, display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 16 }}>
            <div>
              <h1 style={{ fontSize: 18, fontWeight: 700, color: C.text, margin: 0, letterSpacing: "-0.02em" }}>{NAV.find((n) => n.id === page)?.label}</h1>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 3 }}>
                <span style={{ width: 6, height: 6, borderRadius: "50%", background: C.emerald, flexShrink: 0, boxShadow: `0 0 6px ${C.emerald}` }} />
                <span style={{ fontSize: 11, color: C.textMuted }}>{data.length} employees · Privacy-first · All data stays in your browser</span>
              </div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
              <div style={{ fontSize: 11, color: C.textMuted, textAlign: "right" }}>
                <div style={{ fontWeight: 500, color: C.textMid }}>{now}</div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 4, padding: "5px 10px", borderRadius: 7, border: `1px solid ${C.border}`, fontSize: 11, color: C.textMuted }}>
                <Shield size={12} color={C.emerald} /> Secure
              </div>
            </div>
          </div>
          <Uploader onData={setData} count={data.length} />
        </header>

        <div style={{ padding: "22px 24px", flex: 1, overflowY: "auto" }}>
          {page === "overview" && <Overview data={data} />}
          {page === "compensation" && <Compensation data={data} />}
          {page === "directory" && <Directory data={data} />}
        </div>

        <footer style={{ padding: "10px 24px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 11, color: C.textMuted }}>TalentPulse Portal · v2.0 · Privacy-first HR analytics</span>
          <span style={{ fontSize: 11, color: C.textMuted }}>Built with React · SheetJS · Recharts</span>
        </footer>
      </main>
    </div>
  );
}
