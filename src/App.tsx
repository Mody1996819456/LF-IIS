import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { useIsMobile } from "@/hooks/use-mobile";
import ReactDOM from "react-dom";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, LogOut, Plus, Trash2, CheckCircle, Database, Save,
  Clock, Search, Edit3, Download, Loader2,
  X, Upload, FileDown, Building2,
  PieChart, BarChart2, Briefcase, Wifi, WifiOff,
  ShoppingCart, ListOrdered, Leaf, Share2, Printer, ShieldCheck,
  CalendarDays, TrendingUp, Package, Users, User, Lock,
  Sparkles, Brain, AlertTriangle, DollarSign, Target, Award
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  PieChart as RePieChart, Pie, Cell, Legend,
  LineChart, Line, Area, AreaChart
} from "recharts";
import { supabase } from "@/integrations/supabase/client";
import { bootstrapOwner, createManagerAccount } from "@/lib/admin-users.functions";

const formatDate = (dateStr: string) => {
  if (!dateStr) return "-";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
};

const formatDateTime = (dateStr: string) => {
  if (!dateStr) return "-";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  const hours = String(d.getHours()).padStart(2, "0");
  const minutes = String(d.getMinutes()).padStart(2, "0");
  return `${day}-${month}-${year} ${hours}:${minutes}`;
};

const arabicToEnglish = (s: string) =>
    String(s)
      .replace(/[٠١٢٣٤٥٦٧٨٩]/g, (c) => String("٠١٢٣٤٥٦٧٨٩".indexOf(c)))
      .replace(/[۰۱۲۳۴۵۶۷۸۹]/g, (c) => String("۰۱۲۳۴۵۶۷۸۹".indexOf(c)));

const englishToArabic = (s: string | number) => String(s);

const parseImportDate = (value: any): string | null => {
    if (value === null || value === undefined || value === "") return null;
    if (typeof value === "number") {
      if (value <= 0) return null;
      const excelEpoch = new Date(1899, 11, 30);
      const date = new Date(excelEpoch.getTime() + value * 86400000);
      return isNaN(date.getTime()) ? null : date.toISOString().split("T")[0];
    }
    if (value instanceof Date) {
      return isNaN(value.getTime()) ? null : value.toISOString().split("T")[0];
    }
    const str = arabicToEnglish(String(value).trim());
    if (!str) return null;
    if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(str)) {
      const parts = str.split(/[-/]/);
      const iso = `${parts[0]}-${parts[1].padStart(2,"0")}-${parts[2].padStart(2,"0")}`;
      return isNaN(new Date(iso).getTime()) ? null : iso;
    }
    if (/^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/.test(str)) {
      const parts = str.split(/[-/]/);
      const iso = `${parts[2]}-${parts[1].padStart(2,"0")}-${parts[0].padStart(2,"0")}`;
      return isNaN(new Date(iso).getTime()) ? null : iso;
    }
    const fallback = new Date(str);
    return isNaN(fallback.getTime()) ? null : fallback.toISOString().split("T")[0];
};

const departments = ["الإدارة", "المطبخ", "الصيانة", "النظافة", "الأمن", "الموارد البشرية", "المالية", "أخرى"];
const statuses = ["لم يتم", "قيد التنفيذ", "تم التنفيذ", "ملغى"];
const units = ["عدد", "كيس", "علبة", "صندوق", "كيلو", "متر", "ساعة", "أخرى"];
const years = Array.from({ length: 21 }, (_, i) => String(2015 + i));


// ==================== ALL SYSTEM TABLES ====================
const ALL_SYSTEM_TABLES = [
  "admin_affairs_purchases",
  "admin_affairs_summary",
  "admin_affairs_vegetables",
  "admin_affairs_assets",
  "admin_affairs_managers",
  "budget_rows",
  "assets_rows",
  "audit_log",
  "admin_reports"
];

const schemas: Record<string, any> = {
    purchases: {
      tableName: "admin_affairs_purchases",
      title: "طلبات الشراء",
      icon: ShoppingCart,
      uniqueKey: ["system_request_no", "item_name"],
      fields: [
        { key: "admin_request_no", label: "رقم طلب الإدارة", type: "text" },
        { key: "system_request_no", label: "رقم الطلب سيستم", type: "text" },
        { key: "request_date", label: "تاريخ الطلب", type: "date" },
        { key: "item_name", label: "البند", type: "text", required: true },
        { key: "unit", label: "الوحدة", type: "select", options: units },
        { key: "quantity_requested", label: "الكمية المطلوبة", type: "number" },
        { key: "quantity_executed", label: "الكمية المنفذة", type: "number" },
        { key: "quantity_remaining", label: "الكمية المتبقية", type: "number", readonly: true },
        { key: "year", label: "العام", type: "select", options: years },
        { key: "executor", label: "المنفذ", type: "select", options: statuses },
        { key: "requesting_department", label: "القسم الطالب", type: "select", options: departments },
        { key: "receipt_date", label: "تاريخ الاستلام", type: "date" },
        { key: "received_by", label: "المستلم", type: "text" },
        { key: "notes", label: "لزوم", type: "textarea" },
        { key: "remarks", label: "ملاحظات", type: "textarea" },
      ],
      excelColumns: {
        "رقم طلب الادارة": "admin_request_no",
        "رقم الطلب سيستم": "system_request_no",
        "تاريخ الطلب": "request_date",
        "البند": "item_name",
        "الوحده": "unit",
        "الكميه المطلوبه": "quantity_requested",
        "الكميه المنفذه": "quantity_executed",
        "الكميه المتبقيه": "quantity_remaining",
        "العام": "year",
        "المنفذ": "executor",
        "القسم الطالب": "requesting_department",
        "تاريخ الاستلام": "receipt_date",
        "المستلم": "received_by",
        "لزوم": "notes",
        "ملاحظات": "remarks",
      }
    },
    summary: {
      tableName: "admin_affairs_summary",
      title: "إجمالي الطلبات",
      icon: ListOrdered,
      uniqueKey: ["request_number"],
      fields: [
        { key: "request_number", label: "رقم الطلب", type: "text", required: true },
        { key: "request_date", label: "تاريخ الطلب", type: "date" },
        { key: "execution_date", label: "تاريخ التنفيذ", type: "date" },
        { key: "description", label: "وصف طلب الشراء", type: "textarea", required: true },
        { key: "year", label: "العام", type: "select", options: years },
        { key: "status", label: "حالة الطلب", type: "select", options: ["قيد الانتظار", "قيد التنفيذ", "مكتمل", "ملغى"] },
        { key: "remaining_items", label: "عدد البنود المتبقية", type: "text" },
        { key: "remarks", label: "ملاحظات", type: "textarea" },
      ],
      excelColumns: {
        "رقم الطلب": "request_number",
        "تاريخ الطلب": "request_date",
        "تاريخ التنفيذ": "execution_date",
        "وصف طلب الشراء": "description",
        "العام": "year",
        "حالة الطلب": "status",
        "عدد البنود المتبقيه": "remaining_items",
        "ملاحظات": "remarks",
      }
    },
    vegetables: {
      tableName: "admin_affairs_vegetables",
      title: "خضار أسبوعي",
      icon: Leaf,
      uniqueKey: ["system_request_no", "item_name"],
      fields: [
        { key: "admin_request_no", label: "رقم طلب الإدارة", type: "text" },
        { key: "system_request_no", label: "رقم الطلب سيستم", type: "text" },
        { key: "request_date", label: "تاريخ الطلب", type: "date" },
        { key: "item_name", label: "البند", type: "text", required: true },
        { key: "unit", label: "الوحدة", type: "select", options: units },
        { key: "quantity_requested", label: "الكمية المطلوبة", type: "number" },
        { key: "quantity_executed", label: "الكمية المنفذة", type: "number" },
        { key: "quantity_remaining", label: "الكمية المتبقية", type: "number", readonly: true },
        { key: "year", label: "العام", type: "select", options: years },
        { key: "executor", label: "المنفذ", type: "select", options: statuses },
        { key: "requesting_department", label: "القسم الطالب", type: "select", options: departments },
        { key: "receipt_date", label: "تاريخ الاستلام", type: "date" },
        { key: "received_by", label: "المستلم", type: "text" },
        { key: "notes", label: "لزوم", type: "textarea" },
        { key: "remarks", label: "ملاحظات", type: "textarea" },
      ],
      excelColumns: {
        "رقم طلب الادارة": "admin_request_no",
        "رقم الطلب سيستم": "system_request_no",
        "تاريخ الطلب": "request_date",
        "البند": "item_name",
        "الوحده": "unit",
        "الكميه المطلوبه": "quantity_requested",
        "الكميه المنفذه": "quantity_executed",
        "الكميه المتبقيه": "quantity_remaining",
        "العام": "year",
        "المنفذ": "executor",
        "القسم الطالب": "requesting_department",
        "تاريخ الاستلام": "receipt_date",
        "المستلم": "received_by",
        "لزوم": "notes",
        "ملاحظات": "remarks",
      }
    },
    assets: {
      tableName: "admin_affairs_assets",
      title: "الأصول",
      icon: Package,
      uniqueKey: ["purchase_order_no", "item_name"],
      fields: [
        { key: "request_date", label: "التاريخ", type: "date" },
        { key: "item_name", label: "اسم الصنف", type: "text", required: true },
        { key: "unit", label: "الوحدة", type: "select", options: units },
        { key: "quantity_added", label: "كمية الاضافة", type: "number" },
        { key: "supplier", label: "المورد", type: "text" },
        { key: "purchase_order_no", label: "رقم أمر الشراء", type: "text" },
        { key: "requesting_department", label: "الإدارة الطالبة والمستفيدة", type: "select", options: departments },
        { key: "asset_location", label: "مكان الأصل", type: "text" },
        { key: "year", label: "العام", type: "select", options: years },
      ],
      excelColumns: {
        "التاريخ": "request_date",
        "اسم الصنف": "item_name",
        "الوحدة": "unit",
        "كمية الاضافة": "quantity_added",
        "المورد": "supplier",
        "رقم أمر الشراء": "purchase_order_no",
        "الإدارة الطالبة والمستفيدة": "requesting_department",
        "مكان الأصل": "asset_location",
        "العام": "year",
      }
    },
    users: {
      tableName: "admin_affairs_managers",
      title: "المستخدمين",
      icon: Users,
      uniqueKey: ["email"],
      fields: [
        { key: "name", label: "الاسم", type: "text", required: true },
        { key: "email", label: "البريد الإلكتروني", type: "text", required: true },
        { key: "password", label: "كلمة المرور", type: "text", required: true },
        { key: "permissions", label: "الصلاحية", type: "select", options: ["مدير كامل الصلاحيات", "مشاهد فقط"], required: true },
      ],
      excelColumns: {
        "الاسم": "name",
        "البريد الإلكتروني": "email",
        "كلمة المرور": "password",
        "الصلاحية": "permissions",
      }
    }
};

const SortTh = ({ label, field, sortField, sortDir, sortDropdown, onSort, onClear, onToggle, align="center" }: any) => {
  const active = sortField === field;
  const open = sortDropdown === field;
  const thRef = useRef<HTMLTableCellElement>(null);
  const [dropPos, setDropPos] = useState<any>(null);

  const handleToggle = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (!open && thRef.current) {
      const rect = thRef.current.getBoundingClientRect();
      const dropW = 165;
      const centered = rect.left + rect.width / 2;
      const vw = window.innerWidth;
      let left = centered;
      let transform = "translateX(-50%)";
      if (centered - dropW / 2 < 8) {
        left = rect.left; transform = "none";
      } else if (centered + dropW / 2 > vw - 8) {
        left = rect.right; transform = "translateX(-100%)";
      }
      setDropPos({ top: rect.bottom + 4, left, transform });
    }
    onToggle(field);
  };

  const dropdown = open && dropPos ? ReactDOM.createPortal(
    <div
      style={{
        position: "fixed", top: dropPos.top, left: dropPos.left, transform: dropPos.transform,
        background: "white", border: "1px solid #e2e8f0", borderRadius: "10px",
        boxShadow: "0 8px 24px rgba(0,0,0,0.2)", zIndex: 99999, minWidth: "165px",
        overflow: "hidden", direction: "rtl",
      }}
      onMouseDown={e => e.stopPropagation()} onClick={e => e.stopPropagation()}
    >
      <button onMouseDown={e => { e.stopPropagation(); onSort(field, "desc"); }}
        style={{ display:"flex", alignItems:"center", gap:"8px", width:"100%", padding:"10px 14px", background: active && sortDir==="desc" ? "#eef2ff" : "white", border:"none", borderBottom:"1px solid #f1f5f9", cursor:"pointer", fontSize:"13px", fontWeight:"700", color:"#1e293b" }}>
        ↓ من الأعلى للأقل
      </button>
      <button onMouseDown={e => { e.stopPropagation(); onSort(field, "asc"); }}
        style={{ display:"flex", alignItems:"center", gap:"8px", width:"100%", padding:"10px 14px", background: active && sortDir==="asc" ? "#eef2ff" : "white", border:"none", borderBottom: active ? "1px solid #f1f5f9" : "none", cursor:"pointer", fontSize:"13px", fontWeight:"700", color:"#1e293b" }}>
        ↑ من الأقل للأعلى
      </button>
      {active && (
        <button onMouseDown={e => { e.stopPropagation(); onClear(); }}
          style={{ display:"flex", alignItems:"center", gap:"8px", width:"100%", padding:"9px 14px", background:"#fff1f2", border:"none", cursor:"pointer", fontSize:"12px", fontWeight:"700", color:"#dc2626" }}>
          ✕ إلغاء الفرز
        </button>
      )}
    </div>,
    document.body
  ) : null;

  return (
    <th ref={thRef} style={{ padding:"10px 8px", backgroundColor:"transparent", color:"white", fontWeight:"800", textAlign: align as any, position:"relative", whiteSpace: "nowrap", border: "2px solid #4338ca", fontSize:"13px" }}>
      <div style={{ display:"flex", alignItems:"center", justifyContent: align === "center" ? "center" : "flex-start", gap:"6px", cursor:"pointer", userSelect:"none" }} onClick={handleToggle}>
        <span style={{ fontSize:"13px" }}>{label}</span>
        <span style={{ background: active ? "rgba(255, 255, 255, 0.2)" : "transparent", color: active ? "#ffffff" : "#c7d2fe", padding:"3px 5px", borderRadius:"6px", fontSize:"11px", border: active ? "1px solid rgba(255, 255, 255, 0.4)" : "1px solid rgba(199, 210, 254, 0.3)" }}>
          {active ? (sortDir === "desc" ? "↓" : "↑") : "⇅"}
        </span>
      </div>
      {dropdown}
    </th>
  );
};

const ClockWidget = ({ headerMode = false }: { headerMode?: boolean }) => {
   const [time, setTime] = useState(new Date());
   useEffect(() => {
       const timer = setInterval(() => setTime(new Date()), 1000);
       return () => clearInterval(timer);
   }, []);

   if (headerMode) {
       return (
           <div className="flex flex-col items-center justify-center px-4 py-2 bg-white/10 backdrop-blur-md rounded-xl border border-white/20 shadow-[0_4px_30px_rgba(0,0,0,0.1)] text-center transition-all select-none">
               <div className="text-lg font-black text-white font-mono tracking-widest leading-none" dir="ltr">
                   {time.toLocaleTimeString('en-US', { hour12: true })}
               </div>
               <div className="text-[10px] text-indigo-100 font-bold mt-1 leading-none">
                   {time.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}
               </div>
           </div>
       );
   }

   return (
       <div className="flex flex-col items-center justify-center p-3 bg-indigo-50/50 rounded-xl mx-3 mb-3 border border-indigo-100 shadow-sm">
           <div className="text-lg font-black text-indigo-700 font-mono" dir="ltr">
               {time.toLocaleTimeString('en-US', { hour12: true })}
           </div>
           <div className="text-[10px] text-indigo-500 font-bold mt-0.5">
               {time.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}
           </div>
       </div>
   );
};

const quotes = [
  "«النجاح هو الانتقال من فشل إلى فشل دون فقدان الحماس»",
  "«الوقت هو العملة الوحيدة التي تمتلكها، فاحذر أن ينفقها غيرك»",
  "«الإدارة الجيدة هي فن جعل المشاكل مثيرة للاهتمام وحلولها بناءة»",
  "«لا تكن أسهل ما في الحياة ولا أصعب ما فيها، بل كن أصدق ما فيها»",
  "«التخطيط السليم يمنع الأداء الضعيف»",
  "«الجودة ليست صدفة، بل هي دائماً نتيجة جهد ذكي»"
];

const QuoteWidget = () => {
    const [quote, setQuote] = useState(quotes[0]);
    useEffect(() => {
        const interval = setInterval(() => {
            setQuote(quotes[Math.floor(Math.random() * quotes.length)]);
        }, 10000);
        return () => clearInterval(interval);
    }, []);
    return (
        <div className="p-3 mx-3 mb-3 bg-gradient-to-r from-amber-50 to-orange-50 rounded-xl border border-amber-100 text-center relative overflow-hidden shadow-inner">
            <span className="text-amber-600/20 absolute -top-2 -right-2 text-5xl">"</span>
            <p className="text-amber-800 text-xs font-bold leading-relaxed relative z-10 italic">
                {quote}
            </p>
        </div>
    );
};

// ==================== Budget & Assets Analytical Cards (Dashboard) ====================
const BudgetAssetsInsights = ({ budgetRows, assetsNewRows }: { budgetRows: any[]; assetsNewRows: any[] }) => {
  const budgetTotal = budgetRows.reduce((s, r) => s + (Number(r.total_cost) || 0), 0);
  const assetsTotal = assetsNewRows.reduce((s, r) => s + (Number(r.total_cost) || 0), 0);
  const combined = budgetTotal + assetsTotal;
  const forecastNextYear = combined * 1.30;
  const forecastIncrease = forecastNextYear - combined;

  const MONTHS_KEYS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  const MONTHS_LABELS = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];

  const monthlyCombined = MONTHS_KEYS.map((k, i) => {
    const bQty = budgetRows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
    const aQty = assetsNewRows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
    const bAvgPrice = budgetRows.length ? budgetRows.reduce((s, r) => s + (Number(r.price) || 0), 0) / budgetRows.length : 0;
    const aAvgPrice = assetsNewRows.length ? assetsNewRows.reduce((s, r) => s + (Number(r.price) || 0), 0) / assetsNewRows.length : 0;
    const cost = bQty * bAvgPrice + aQty * aAvgPrice;
    return { name: MONTHS_LABELS[i], "التكلفة": Math.round(cost), "المتوقع": Math.round(cost * 1.30) };
  });

  const topBudgetItems = [...budgetRows]
    .sort((a, b) => (Number(b.total_cost) || 0) - (Number(a.total_cost) || 0))
    .slice(0, 5);

  const topAssetsItems = [...assetsNewRows]
    .sort((a, b) => (Number(b.total_cost) || 0) - (Number(a.total_cost) || 0))
    .slice(0, 5);

  const distributionData = [
    { name: "الموازنة التشغيلية", value: budgetTotal, color: "#f59e0b" },
    { name: "الأصول الجديدة", value: assetsTotal, color: "#8b5cf6" },
  ].filter(d => d.value > 0);

  const sheetTotals: Record<string, number> = {};
  budgetRows.forEach(r => {
    const sheet = r.sheet || "غير محدد";
    sheetTotals[sheet] = (sheetTotals[sheet] || 0) + (Number(r.total_cost) || 0);
  });
  const sheetsData = Object.entries(sheetTotals)
    .map(([name, cost]) => ({ name: name.length > 18 ? name.slice(0, 15) + "…" : name, cost }))
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 6);

  const deptTotals: Record<string, number> = {};
  assetsNewRows.forEach(r => {
    const dept = r.department || "غير محدد";
    deptTotals[dept] = (deptTotals[dept] || 0) + (Number(r.total_cost) || 0);
  });
  const deptData = Object.entries(deptTotals)
    .map(([name, cost]) => ({ name, cost }))
    .sort((a, b) => b.cost - a.cost)
    .slice(0, 6);

  const avgItemCost = combined && (budgetRows.length + assetsNewRows.length) > 0
    ? Math.round(combined / (budgetRows.length + assetsNewRows.length))
    : 0;

  const peakMonth = [...monthlyCombined].sort((a, b) => b["التكلفة"] - a["التكلفة"])[0];

  if (combined === 0) return null;

  return (
    <div className="space-y-3" dir="rtl">
      <div className="bg-gradient-to-l from-teal-600 via-cyan-600 to-sky-600 text-white p-3 rounded-xl shadow-sm flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-2">
          <div className="p-1.5 bg-white/20 backdrop-blur rounded-lg">
            <Brain size={16} />
          </div>
          <div>
            <h3 className="text-sm font-black m-0">تحليلات الموازنة والأصول 💎</h3>
            <p className="text-white/90 text-[10px] font-bold m-0">توقعات العام القادم بزيادة 30%</p>
          </div>
        </div>
        <div className="bg-white/15 backdrop-blur rounded-lg px-3 py-1.5 border border-white/30">
          <p className="text-[9px] font-bold text-white/80 m-0">الإجمالي الشامل</p>
          <p className="text-sm font-black m-0">{combined.toLocaleString("en-US")} ج</p>
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"8px" }}>
        {/* Card 1 — إجمالي الموازنة */}
        <div style={{ background:"linear-gradient(135deg,#ecfeff,#cffafe)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #22d3ee", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#06b6d4", borderRadius:"8px", display:"flex" }}><DollarSign size={13} color="white" /></div>
            <BarChart2 size={12} style={{ color:"#0891b2", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#164e63" }}>إجمالي الموازنة</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#0e7490" }}>{budgetTotal.toLocaleString("en-US")} ج</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#0891b2" }}>{budgetRows.length} صنف</p>
        </div>

        {/* Card 2 — إجمالي الأصول */}
        <div style={{ background:"linear-gradient(135deg,#f0fdfa,#ccfbf1)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #2dd4bf", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#14b8a6", borderRadius:"8px", display:"flex" }}><Building2 size={13} color="white" /></div>
            <Package size={12} style={{ color:"#0d9488", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#134e4a" }}>إجمالي الأصول الجديدة</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#0f766e" }}>{assetsTotal.toLocaleString("en-US")} ج</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#0d9488" }}>{assetsNewRows.length} أصل</p>
        </div>

        {/* Card 3 — توقع العام القادم */}
        <div style={{ background:"linear-gradient(135deg,#eff6ff,#dbeafe)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #60a5fa", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#3b82f6", borderRadius:"8px", display:"flex" }}><TrendingUp size={13} color="white" /></div>
            <Sparkles size={12} style={{ color:"#2563eb", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#1e3a8a" }}>توقع العام القادم (+30%)</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#1d4ed8" }}>{forecastNextYear.toLocaleString("en-US")} ج</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#2563eb" }}>+{forecastIncrease.toLocaleString("en-US")} ج زيادة</p>
        </div>

        {/* Card 4 — متوسط تكلفة الصنف */}
        <div style={{ background:"linear-gradient(135deg,#ecfdf5,#d1fae5)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #34d399", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#10b981", borderRadius:"8px", display:"flex" }}><Target size={13} color="white" /></div>
            <Award size={12} style={{ color:"#059669", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#064e3b" }}>متوسط تكلفة الصنف</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#059669" }}>{avgItemCost.toLocaleString("en-US")} ج</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#10b981" }}>ذروة: {peakMonth?.name || "-"}</p>
        </div>
      </div>

      <div className="bg-white p-3 rounded-xl border shadow-sm">
        <h3 className="text-[11px] font-black mb-2 text-purple-900 bg-purple-100/60 p-2 rounded-lg flex items-center gap-2">
          <TrendingUp size={14} className="text-purple-600" />
          توقع التكاليف الشهرية للعام القادم (زيادة 30%)
        </h3>
        <div className="h-56">
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={monthlyCombined} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
              <defs>
                <linearGradient id="curColor" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.7} />
                  <stop offset="95%" stopColor="#3b82f6" stopOpacity={0.05} />
                </linearGradient>
                <linearGradient id="futColor" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.7} />
                  <stop offset="95%" stopColor="#f59e0b" stopOpacity={0.05} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 10, fill: "#64748b" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 10, fill: "#64748b" }} axisLine={false} tickLine={false} />
              <RechartsTooltip
                formatter={(v: any) => Number(v).toLocaleString("en-US") + " ج"}
                contentStyle={{ borderRadius: "12px", border: "none", boxShadow: "0 4px 20px rgba(0,0,0,0.1)" }}
              />
              <Legend wrapperStyle={{ fontSize: "11px", fontWeight: 700 }} />
              <Area type="monotone" dataKey="التكلفة" stroke="#3b82f6" fill="url(#curColor)" strokeWidth={2} />
              <Area type="monotone" dataKey="المتوقع" stroke="#f59e0b" fill="url(#futColor)" strokeWidth={2} strokeDasharray="5 5" />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="bg-white p-3 rounded-xl border shadow-sm">
          <h3 className="text-[11px] font-black mb-2 text-purple-900 bg-purple-100/60 p-2 rounded-lg flex items-center gap-2">
            <PieChart size={14} className="text-purple-600" />
            توزيع التكلفة بين الأقسام
          </h3>
          <div className="h-52">
            <ResponsiveContainer width="100%" height="100%">
              <RePieChart>
                <Pie data={distributionData} cx="50%" cy="50%" innerRadius={50} outerRadius={80} dataKey="value" paddingAngle={4}>
                  {distributionData.map((e, i) => (<Cell key={i} fill={e.color} />))}
                </Pie>
                <RechartsTooltip formatter={(v: any) => Number(v).toLocaleString("en-US") + " ج"} contentStyle={{ borderRadius: "10px", border: "none", boxShadow: "0 4px 20px rgba(0,0,0,0.1)" }} />
                <Legend wrapperStyle={{ fontSize: "11px", fontWeight: 700 }} />
              </RePieChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white p-3 rounded-xl border shadow-sm">
          <h3 className="text-[11px] font-black mb-2 text-purple-900 bg-purple-100/60 p-2 rounded-lg flex items-center gap-2">
            <BarChart2 size={14} className="text-purple-600" />
            أعلى شيتات الموازنة تكلفة
          </h3>
          <div className="h-52">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={sheetsData} layout="vertical" margin={{ top: 5, right: 20, left: 60, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 9, fill: "#64748b" }} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 10, fill: "#334155", fontWeight: 700 }} axisLine={false} tickLine={false} width={80} />
                <RechartsTooltip formatter={(v: any) => Number(v).toLocaleString("en-US") + " ج"} contentStyle={{ borderRadius: "10px", border: "none", boxShadow: "0 4px 20px rgba(0,0,0,0.1)" }} />
                <Bar dataKey="cost" radius={[0, 8, 8, 0]} barSize={18} fill="#f59e0b" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="bg-white p-3 rounded-xl border shadow-sm">
          <h3 className="text-[11px] font-black mb-2 text-amber-900 bg-amber-100/60 p-2 rounded-lg flex items-center gap-2">
            <Award size={14} className="text-amber-600" />
            أعلى 5 أصناف في الموازنة
          </h3>
          {topBudgetItems.length === 0 ? (
            <p className="text-slate-400 text-center py-6 font-bold text-sm">لا توجد بيانات</p>
          ) : (
            <div className="space-y-2">
              {topBudgetItems.map((it, i) => {
                const pct = budgetTotal > 0 ? ((Number(it.total_cost) || 0) / budgetTotal) * 100 : 0;
                return (
                  <div key={it.id || i} className="p-2 rounded-lg bg-amber-50/60 border border-amber-100">
                    <div className="flex items-center justify-between mb-1">
                      <span className="font-black text-slate-800 text-xs truncate flex-1">{i + 1}. {it.item}</span>
                      <span className="font-black text-amber-700 text-xs mr-2 whitespace-nowrap">{Number(it.total_cost).toLocaleString("en-US")} ج</span>
                    </div>
                    <div className="w-full bg-amber-100 rounded-full h-1.5 overflow-hidden">
                      <div className="h-1.5 rounded-full bg-gradient-to-l from-amber-500 to-orange-500" style={{ width: `${pct}%` }} />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="bg-white p-3 rounded-xl border shadow-sm">
          <h3 className="text-[11px] font-black mb-2 text-violet-900 bg-violet-100/60 p-2 rounded-lg flex items-center gap-2">
            <Award size={14} className="text-violet-600" />
            أعلى 5 أصول تكلفة
          </h3>
          {topAssetsItems.length === 0 ? (
            <p className="text-slate-400 text-center py-6 font-bold text-sm">لا توجد بيانات</p>
          ) : (
            <div className="space-y-2">
              {topAssetsItems.map((it, i) => {
                const pct = assetsTotal > 0 ? ((Number(it.total_cost) || 0) / assetsTotal) * 100 : 0;
                return (
                  <div key={it.id || i} className="p-2 rounded-lg bg-violet-50/60 border border-violet-100">
                    <div className="flex items-center justify-between mb-1">
                      <span className="font-black text-slate-800 text-xs truncate flex-1">{i + 1}. {it.item_name}</span>
                      <span className="font-black text-violet-700 text-xs mr-2 whitespace-nowrap">{Number(it.total_cost).toLocaleString("en-US")} ج</span>
                    </div>
                    <div className="w-full bg-violet-100 rounded-full h-1.5 overflow-hidden">
                      <div className="h-1.5 rounded-full bg-gradient-to-l from-violet-500 to-purple-500" style={{ width: `${pct}%` }} />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {deptData.length > 0 && (
        <div className="bg-white p-3 rounded-xl border shadow-sm">
          <h3 className="text-[11px] font-black mb-2 text-purple-900 bg-purple-100/60 p-2 rounded-lg flex items-center gap-2">
            <Building2 size={14} className="text-purple-600" />
            توزيع تكلفة الأصول حسب القسم
          </h3>
          <div className="h-52">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={deptData} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b", fontWeight: 700 }} axisLine={false} tickLine={false} />
                <YAxis tick={{ fontSize: 9, fill: "#64748b" }} axisLine={false} tickLine={false} />
                <RechartsTooltip formatter={(v: any) => Number(v).toLocaleString("en-US") + " ج"} contentStyle={{ borderRadius: "10px", border: "none", boxShadow: "0 4px 20px rgba(0,0,0,0.1)" }} />
                <Bar dataKey="cost" radius={[8, 8, 0, 0]} barSize={32} fill="#8b5cf6" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}
    </div>
  );
};


// ==================== ADMIN REPORTS INSIGHTS (Dashboard) ====================
const AdminReportsInsights = ({ adminReports }: { adminReports: any[] }) => {
  const PIE_COLORS = ["#6366f1","#0d9488","#f59e0b","#f43f5e","#8b5cf6","#14b8a6","#3b82f6","#ec4899","#10b981","#ef4444","#a855f7","#06b6d4"];

  const totalValue = adminReports.reduce((s, r) => s + (Number(r.total_value) || 0), 0);
  const totalQty   = adminReports.reduce((s, r) => s + (Number(r.quantity) || 0), 0);
  const uniqueStores = new Set(adminReports.map(r => r.store_name).filter(Boolean)).size;
  const uniqueTasks  = new Set(adminReports.map(r => r.task).filter(Boolean)).size;

  // Monthly spending — based on report_month field
  const MONTHS_ORDER = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];
  const monthlyData = MONTHS_ORDER.map(m => {
    const val = adminReports.filter(r => r.report_month === m).reduce((s, r) => s + (Number(r.total_value) || 0), 0);
    return { name: m.slice(0, 3), fullName: m, value: Math.round(val) };
  }).filter(m => m.value > 0);

  const peakMonth = monthlyData.length ? [...monthlyData].sort((a,b) => b.value - a.value)[0] : null;

  // Top stores
  const storeMap: Record<string,number> = {};
  adminReports.forEach(r => { const k = r.store_name || "غير محدد"; storeMap[k] = (storeMap[k]||0) + (Number(r.total_value)||0); });
  const topStores = Object.entries(storeMap).map(([name,value]) => ({ name, value: Math.round(value) })).sort((a,b) => b.value - a.value).slice(0,6);

  // Top tasks
  const taskMap: Record<string,number> = {};
  adminReports.forEach(r => { const k = r.task_description || r.task || "غير محدد"; taskMap[k] = (taskMap[k]||0) + (Number(r.total_value)||0); });
  const topTasks = Object.entries(taskMap).map(([name,value]) => ({ name: name.length > 18 ? name.slice(0,16)+"…" : name, value: Math.round(value) })).sort((a,b) => b.value - a.value).slice(0,6);

  if (adminReports.length === 0) return null;

  return (
    <div style={{ background:"white", borderRadius:"16px", border:"1px solid #e2e8f0", padding:"16px", boxShadow:"0 4px 20px rgba(0,0,0,0.06)" }}>
      {/* Header */}
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"12px", flexWrap:"wrap", gap:"8px" }}>
        <h3 style={{ margin:0, fontWeight:"900", fontSize:"14px", color:"#0f766e", display:"flex", alignItems:"center", gap:"6px" }}>
          <CalendarDays size={16} /> تحليلات التقارير الإدارية 📋
        </h3>
        <span style={{ background:"#f0fdfa", color:"#0f766e", padding:"3px 10px", borderRadius:"12px", fontSize:"11px", fontWeight:"800", border:"1px solid #5eead4" }}>
          {adminReports.length.toLocaleString("en-US")} سجل
        </span>
      </div>

      {/* KPI Cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"8px", marginBottom:"14px" }}>
        {/* Card 1 — إجمالي المصروفات */}
        <div style={{ background:"linear-gradient(135deg,#f0fdfa,#ccfbf1)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #2dd4bf", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#14b8a6", borderRadius:"8px", display:"flex" }}><DollarSign size={13} color="white" /></div>
            <BarChart2 size={12} style={{ color:"#0d9488", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#134e4a" }}>إجمالي المصروفات</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#0f766e" }}>{totalValue.toLocaleString("en-US", { maximumFractionDigits:0 })} ج</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#0d9488" }}>{adminReports.length} سجل</p>
        </div>
        {/* Card 2 — إجمالي الكميات */}
        <div style={{ background:"linear-gradient(135deg,#fffbeb,#fef3c7)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #fcd34d", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#f59e0b", borderRadius:"8px", display:"flex" }}><Package size={13} color="white" /></div>
            <TrendingUp size={12} style={{ color:"#d97706", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#78350f" }}>إجمالي الكميات</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#b45309" }}>{totalQty.toLocaleString("en-US")}</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#d97706" }}>وحدة مصروفة</p>
        </div>
        {/* Card 3 — عدد المخازن */}
        <div style={{ background:"linear-gradient(135deg,#f5f3ff,#ede9fe)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #c4b5fd", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#8b5cf6", borderRadius:"8px", display:"flex" }}><Building2 size={13} color="white" /></div>
            <Users size={12} style={{ color:"#6d28d9", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#3b0764" }}>عدد المخازن</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#6d28d9" }}>{uniqueStores}</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#7c3aed" }}>{uniqueTasks} task مختلفة</p>
        </div>
        {/* Card 4 — ذروة الصرف */}
        <div style={{ background:"linear-gradient(135deg,#eff6ff,#dbeafe)", padding:"10px 12px", borderRadius:"10px", border:"1px solid #60a5fa", position:"relative", overflow:"hidden" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"6px" }}>
            <div style={{ padding:"5px", background:"#3b82f6", borderRadius:"8px", display:"flex" }}><CalendarDays size={13} color="white" /></div>
            <Award size={12} style={{ color:"#2563eb", opacity:0.35 }} />
          </div>
          <p style={{ margin:0, fontSize:"10px", fontWeight:"800", color:"#1e3a8a" }}>ذروة الصرف</p>
          <p style={{ margin:"4px 0 0", fontSize:"15px", fontWeight:"900", color:"#1d4ed8" }}>{peakMonth?.fullName || "-"}</p>
          <p style={{ margin:"3px 0 0", fontSize:"9px", fontWeight:"700", color:"#3b82f6" }}>{peakMonth ? peakMonth.value.toLocaleString("en-US") + " ج" : "لا بيانات"}</p>
        </div>
      </div>

      {/* Charts Row */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:"12px" }}>

        {/* Monthly Bar Chart */}
        <div style={{ background:"#f8fafc", borderRadius:"10px", padding:"10px", border:"1px solid #e2e8f0" }}>
          <p style={{ margin:"0 0 6px", fontWeight:"800", fontSize:"11px", color:"#0f766e" }}>📅 المصروفات الشهرية</p>
          <div style={{ height:"160px" }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={monthlyData} margin={{ top:0, right:0, left:-22, bottom:0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" tick={{ fontSize:8, fill:"#64748b" }} axisLine={false} tickLine={false} />
                <YAxis tick={{ fontSize:8, fill:"#64748b" }} axisLine={false} tickLine={false} tickFormatter={v => v >= 1000 ? (v/1000).toFixed(0)+"k" : v} />
                <RechartsTooltip
                  formatter={(v: any) => [Number(v).toLocaleString("en-US") + " ج", "القيمة"]}
                  labelFormatter={(_,p) => p?.[0]?.payload?.fullName || ""}
                  contentStyle={{ borderRadius:"8px", border:"none", boxShadow:"0 4px 20px rgba(0,0,0,0.1)", fontSize:"11px" }}
                />
                <Bar dataKey="value" radius={[6,6,0,0]} barSize={14}>
                  {monthlyData.map((_,i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Top Stores */}
        <div style={{ background:"#f8fafc", borderRadius:"10px", padding:"10px", border:"1px solid #e2e8f0" }}>
          <p style={{ margin:"0 0 8px", fontWeight:"800", fontSize:"11px", color:"#0f766e" }}>🏪 أعلى المخازن إنفاقاً</p>
          <div style={{ display:"flex", flexDirection:"column", gap:"5px" }}>
            {topStores.map((s,i) => {
              const pct = totalValue > 0 ? Math.round((s.value / totalValue) * 100) : 0;
              return (
                <div key={i}>
                  <div style={{ display:"flex", justifyContent:"space-between", marginBottom:"2px" }}>
                    <span style={{ fontSize:"9px", color:"#64748b", fontWeight:"700" }}>{pct}%</span>
                    <span style={{ fontSize:"10px", fontWeight:"700", color:"#1e293b", maxWidth:"130px", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{s.name}</span>
                  </div>
                  <div style={{ background:"#e2e8f0", borderRadius:"99px", height:"5px" }}>
                    <div style={{ height:"100%", width:`${pct}%`, background:PIE_COLORS[i % PIE_COLORS.length], borderRadius:"99px" }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Top Tasks Pie */}
        <div style={{ background:"#f8fafc", borderRadius:"10px", padding:"10px", border:"1px solid #e2e8f0" }}>
          <p style={{ margin:"0 0 6px", fontWeight:"800", fontSize:"11px", color:"#0f766e" }}>🎯 توزيع المصروفات (Task)</p>
          <div style={{ height:"160px" }}>
            <ResponsiveContainer width="100%" height="100%">
              <RePieChart>
                <Pie data={topTasks} cx="50%" cy="50%" innerRadius={35} outerRadius={60} dataKey="value" paddingAngle={3}>
                  {topTasks.map((_,i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                </Pie>
                <RechartsTooltip
                  formatter={(v: any) => [Number(v).toLocaleString("en-US") + " ج", "القيمة"]}
                  contentStyle={{ borderRadius:"8px", border:"none", boxShadow:"0 4px 20px rgba(0,0,0,0.1)", fontSize:"11px" }}
                />
                <Legend wrapperStyle={{ fontSize:"9px", fontWeight:700 }} />
              </RePieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );
};

// ─── Supabase pagination helper ───────────────────────────────────────────────
// Supabase defaults to 1000 rows per request. This helper loops using .range()
// until every row is fetched, regardless of total count.
const fetchAllPages = async (makeQuery: () => any): Promise<any[]> => {
  const PAGE = 1000;
  let all: any[] = [];
  let from = 0;
  while (true) {
    const { data, error } = await makeQuery().range(from, from + PAGE - 1);
    if (error) throw error;
    all = all.concat(data || []);
    if (!data || data.length < PAGE) break;
    from += PAGE;
  }
  return all;
};
// ──────────────────────────────────────────────────────────────────────────────

const Dashboard = ({ supabase, systemMenu }: { supabase: any, systemMenu?: React.ReactNode }) => {
   const [stats, setStats] = useState({ purchases: [], summary: [], vegetables: [], assets: [], budget: [], assets_new: [], admin_reports: [] });
   const [loading, setLoading] = useState(true);

   useEffect(() => {
       const fetchStats = async () => {
           setLoading(true);
           const [purchases_d, summary_d, vegetables_d, assets_d, budget_d, assetsNew_d, adminRep_d] = await Promise.all([
               fetchAllPages(() => supabase.from("admin_affairs_purchases").select("*").order('request_date', { ascending: false })),
               fetchAllPages(() => supabase.from("admin_affairs_summary").select("*")),
               fetchAllPages(() => supabase.from("admin_affairs_vegetables").select("*")),
               fetchAllPages(() => supabase.from("admin_affairs_assets").select("*")),
               fetchAllPages(() => supabase.from("budget_rows").select("*")),
               fetchAllPages(() => supabase.from("assets_rows").select("*")),
               fetchAllPages(() => supabase.from("admin_reports").select("*"))
           ]);
           setStats({ purchases: purchases_d, summary: summary_d, vegetables: vegetables_d, assets: assets_d, budget: budget_d, assets_new: assetsNew_d, admin_reports: adminRep_d });
           setLoading(false);
       };
       fetchStats();
   }, [supabase]);

   if (loading) return <div className="flex items-center justify-center h-64"><Loader2 className="animate-spin text-indigo-600" size={40} /></div>;

   const purchases = stats.purchases as any[];
   const vegetables = stats.vegetables as any[];
   const summary = stats.summary as any[];
   const assets = stats.assets as any[];
   const budgetRows = stats.budget as any[];
   const adminReports = (stats as any).admin_reports as any[];
   const assetsNewRows = stats.assets_new as any[];
   const budgetTotalCost = budgetRows.reduce((s: number, r: any) => s + (Number(r.total_cost) || 0), 0);
   const assetsNewTotalCost = assetsNewRows.reduce((s: number, r: any) => s + (Number(r.total_cost) || 0), 0);

   const isCompleted = (val: any, type: 'executor' | 'status' = 'executor') => {
       if (!val) return false;
       const v = String(val).trim();
       if (type === 'status') {
           return v === "مكتمل" || v === "تم التنفيذ" || v === "تم";
       }
       return v === "تم التنفيذ" || v === "مكتمل" || v === "تم";
   };

   const completedPurchases = purchases.filter(p => isCompleted(p.executor, 'executor')).length;
   const pendingPurchases = purchases.length - completedPurchases;

   const completedVeg = vegetables.filter(v => isCompleted(v.executor, 'executor')).length;

   const completedSummary = summary.filter(s => isCompleted(s.status, 'status')).length;
   const pendingSummary = summary.length - completedSummary;

   const completedAssets = assets.length;

   const totalActionable = purchases.length + vegetables.length + summary.length + assets.length;
   const totalCompleted = completedPurchases + completedVeg + completedSummary + completedAssets;
   const completionRate = totalActionable > 0 ? Math.round((totalCompleted / totalActionable) * 100) : 0;
   const purCompletionRate = purchases.length > 0 ? Math.round((completedPurchases / purchases.length) * 100) : 0;
   const vegCompletionRate = vegetables.length > 0 ? Math.round((completedVeg / vegetables.length) * 100) : 0;
   const sumCompletionRate = summary.length > 0 ? Math.round((completedSummary / summary.length) * 100) : 0;

   const deptData = departments.map(d => ({
       name: d,
       count: purchases.filter(p => p.requesting_department === d).length + assets.filter(a => a.requesting_department === d).length
   })).filter(d => d.count > 0);

   const yearsCount: Record<string, number> = {};
   [...purchases, ...vegetables, ...summary, ...assets].forEach(r => {
       const y = r.year || (r.request_date ? new Date(r.request_date).getFullYear().toString() : "غير محدد");
       yearsCount[y] = (yearsCount[y] || 0) + 1;
   });
   const yearData = Object.keys(yearsCount).map(y => ({ name: y, count: yearsCount[y] })).sort((a,b) => a.name.localeCompare(b.name));

   const latestPurchases = purchases.slice(0, 5);

   return (
       <div className="space-y-4 min-h-screen pb-6" style={{ fontFamily: "Cairo, sans-serif" }}>
           <div className="mb-4 w-full">
               <table className="w-full border-collapse shadow-sm rounded-xl overflow-visible relative z-10">
                   <tbody>
                       <tr>
                           <td className="bg-gradient-to-r from-indigo-600 to-blue-700 text-white px-6 py-3 w-full relative overflow-visible rounded-xl">
                               <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 w-full">
                                   <div className="flex flex-col md:flex-row md:items-center justify-start gap-4">
                                       <h2 className="text-2xl font-black flex items-center gap-2 m-0">
                                           لوحة التحليلات 📊
                                       </h2>
                                       <p className="text-indigo-100 text-xs m-0 font-bold border-r border-indigo-400/50 pr-4">نظرة شاملة على بيانات نظام المشتريات والشؤون الإدارية</p>
                                   </div>
                                   <div className="flex items-center gap-3 mr-auto md:mr-0 md:mr-auto justify-end relative z-[99]">
                                       <img 
                                           src="/logo.png" 
                                           alt="Linah Farms" 
                                           className="h-14 w-auto object-contain select-none"
                                           style={{ border: "none", outline: "none", boxShadow: "none" }}
                                           referrerPolicy="no-referrer"
                                           onError={(e) => { e.currentTarget.style.display = 'none'; }} 
                                       />
                                       <ClockWidget headerMode={true} />
                                       {systemMenu}
                                   </div>
                               </div>
                           </td>
                       </tr>
                   </tbody>
               </table>
           </div>

           <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-2">
               {/* طلبات الشراء */}
               <div className="bg-blue-50/40 p-3 rounded-xl border border-blue-200 shadow-sm text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-1 h-full bg-blue-500 rounded-r-xl"></div>
                  <div className="flex justify-between items-center mb-1.5">
                      <div className="p-1.5 bg-blue-100/80 text-blue-600 rounded-lg"><ShoppingCart size={16} /></div>
                      <p className="text-blue-950 font-black text-[11px]">طلبات الشراء</p>
                  </div>
                  <h3 className="text-xl font-black text-blue-600 my-1">{englishToArabic(purchases.length)}</h3>
                  <p className="text-[10px] text-blue-800 font-black">{englishToArabic(completedPurchases)} منجز · {englishToArabic(pendingPurchases)} معلق</p>
               </div>

               {/* إجمالي الطلبات */}
               <div className="bg-purple-50/40 p-3 rounded-xl border border-purple-200 shadow-sm text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-1 h-full bg-purple-500 rounded-r-xl"></div>
                  <div className="flex justify-between items-center mb-1.5">
                      <div className="p-1.5 bg-purple-100/80 text-purple-600 rounded-lg"><ListOrdered size={16} /></div>
                      <p className="text-purple-950 font-black text-[11px]">إجمالي الطلبات</p>
                  </div>
                  <h3 className="text-xl font-black text-purple-600 my-1">{englishToArabic(summary.length)}</h3>
                  <p className="text-[10px] text-purple-800 font-black">{englishToArabic(completedSummary)} منجز · {englishToArabic(pendingSummary)} معلق</p>
               </div>

               {/* الخضار الأسبوعي */}
               <div className="bg-green-50/40 p-3 rounded-xl border border-green-200 shadow-sm text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-1 h-full bg-green-500 rounded-r-xl"></div>
                  <div className="flex justify-between items-center mb-1.5">
                      <div className="p-1.5 bg-green-100/80 text-green-600 rounded-lg"><Leaf size={16} /></div>
                      <p className="text-green-950 font-black text-[11px]">الخضار الأسبوعي</p>
                  </div>
                  <h3 className="text-xl font-black text-green-600 my-1">{englishToArabic(vegetables.length)}</h3>
                  <p className="text-[10px] text-green-800 font-black">{englishToArabic(completedVeg)} منجز</p>
               </div>

               {/* الأصول */}
               <div className="bg-orange-50/40 p-3 rounded-xl border border-orange-200 shadow-sm text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-1 h-full bg-orange-500 rounded-r-xl"></div>
                  <div className="flex justify-between items-center mb-1.5">
                      <div className="p-1.5 bg-orange-100/80 text-orange-600 rounded-lg"><Package size={16} /></div>
                      <p className="text-orange-950 font-black text-[11px]">الأصول</p>
                  </div>
                  <h3 className="text-xl font-black text-orange-600 my-1">{englishToArabic(assets.length)}</h3>
                  <p className="text-[10px] text-orange-800 font-black">{englishToArabic(completedAssets)} إجمالي</p>
               </div>

               {/* نسبة الإنجاز */}
               <div className="bg-emerald-50/40 p-3 rounded-xl border border-emerald-200 shadow-sm text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-1 h-full bg-emerald-500 rounded-r-xl"></div>
                  <div className="flex justify-between items-center mb-1.5">
                      <div className="p-1.5 bg-emerald-100/80 text-emerald-600 rounded-lg"><CheckCircle size={16} /></div>
                      <p className="text-emerald-950 font-black text-[11px]">نسبة الإنجاز</p>
                  </div>
                  <h3 className="text-xl font-black text-emerald-600 my-1">{englishToArabic(completionRate)}%</h3>
                  <p className="text-[10px] text-emerald-800 font-black">{englishToArabic(totalCompleted)} من {englishToArabic(totalActionable)}</p>
               </div>

               {/* الموازنة */}
               <div className="bg-amber-50/40 p-3 rounded-xl border border-amber-200 shadow-sm text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-1 h-full bg-amber-500 rounded-r-xl"></div>
                  <div className="flex justify-between items-center mb-1.5">
                      <div className="p-1.5 bg-amber-100/80 text-amber-600 rounded-lg"><BarChart2 size={16} /></div>
                      <p className="text-amber-950 font-black text-[11px]">الموازنة</p>
                  </div>
                  <h3 className="text-xl font-black text-amber-600 my-1">{englishToArabic(budgetRows.length)}</h3>
                  <p className="text-[10px] text-amber-800 font-black">{budgetTotalCost.toLocaleString("en-US")} ج</p>
               </div>

               {/* الأصول الجديدة */}
               <div className="bg-violet-50/40 p-3 rounded-xl border border-violet-200 shadow-sm text-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-1 h-full bg-violet-500 rounded-r-xl"></div>
                  <div className="flex justify-between items-center mb-1.5">
                      <div className="p-1.5 bg-violet-100/80 text-violet-600 rounded-lg"><Building2 size={16} /></div>
                      <p className="text-violet-950 font-black text-[11px]">الأصول الجديدة</p>
                  </div>
                  <h3 className="text-xl font-black text-violet-600 my-1">{englishToArabic(assetsNewRows.length)}</h3>
                  <p className="text-[10px] text-violet-800 font-black">{assetsNewTotalCost.toLocaleString("en-US")} ج</p>
               </div>
            </div>

           <div className="bg-white p-4 rounded-xl border shadow-sm">
               <h3 className="text-sm font-bold mb-3 text-purple-900 bg-purple-100/60 p-2 rounded-xl flex items-center gap-2">
                   <TrendingUp size={16} className="text-purple-600" /> معدل الإنجاز
               </h3>
               <div className="space-y-3">
                   <div>
                       <div className="flex justify-between text-xs mb-1">
                           <span className="text-indigo-600 font-bold">{englishToArabic(completedPurchases)} / {englishToArabic(purchases.length)} ({englishToArabic(purCompletionRate)}%)</span>
                           <span className="font-bold text-slate-700">طلبات الشراء</span>
                       </div>
                       <div className="w-full bg-slate-100 rounded-full h-2 flex justify-end">
                           <div className="bg-indigo-600 h-2 rounded-full" style={{ width: `${purCompletionRate}%` }}></div>
                       </div>
                   </div>
                   <div>
                       <div className="flex justify-between text-xs mb-1">
                           <span className="text-purple-600 font-bold">{englishToArabic(completedSummary)} / {englishToArabic(summary.length)} ({englishToArabic(sumCompletionRate)}%)</span>
                           <span className="font-bold text-slate-700">إجمالي الطلبات</span>
                       </div>
                       <div className="w-full bg-slate-100 rounded-full h-2 flex justify-end">
                           <div className="bg-purple-500 h-2 rounded-full" style={{ width: `${sumCompletionRate}%` }}></div>
                       </div>
                   </div>
                   <div>
                       <div className="flex justify-between text-xs mb-1">
                           <span className="text-emerald-600 font-bold">{englishToArabic(completedVeg)} / {englishToArabic(vegetables.length)} ({englishToArabic(vegCompletionRate)}%)</span>
                           <span className="font-bold text-slate-700">الخضار الأسبوعي</span>
                       </div>
                       <div className="w-full bg-slate-100 rounded-full h-2 flex justify-end">
                           <div className="bg-emerald-500 h-2 rounded-full" style={{ width: `${vegCompletionRate}%` }}></div>
                       </div>
                   </div>
                   <div>
                       <div className="flex justify-between text-xs mb-1">
                           <span className="text-orange-500 font-bold">{englishToArabic(completedAssets)} / {englishToArabic(assets.length)} ({englishToArabic(assets.length > 0 ? 100 : 0)}%)</span>
                           <span className="font-bold text-slate-700">الأصول</span>
                       </div>
                       <div className="w-full bg-slate-100 rounded-full h-2 flex justify-end">
                           <div className="bg-orange-500 h-2 rounded-full" style={{ width: `${assets.length > 0 ? 100 : 0}%` }}></div>
                       </div>
                   </div>
               </div>
           </div>

           <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
               <div className="bg-white p-4 rounded-xl border shadow-sm flex flex-col">
                  <h3 className="text-xs font-bold mb-3 text-purple-900 bg-purple-100/60 p-2 rounded-xl text-center">حالة جميع السجلات 🟢</h3>
                  <div className="flex-1 flex flex-col justify-center">
                      <div className="flex justify-between items-center mb-4 px-2">
                          <div className="text-right">
                              <p className="text-emerald-600 font-bold mb-1 text-xs">✅ تم التنفيذ</p>
                              <p className="text-2xl font-black text-slate-800">{englishToArabic(totalCompleted)} <span className="text-[10px] text-slate-400 font-normal">سجل</span></p>
                          </div>
                          <div className="text-left">
                              <p className="text-red-500 font-bold mb-1 text-xs">معلق / لم يتم ⏳</p>
                              <p className="text-2xl font-black text-slate-800">{englishToArabic(totalActionable - totalCompleted)} <span className="text-[10px] text-slate-400 font-normal">سجل</span></p>
                          </div>
                      </div>

                      <div className="w-full bg-slate-100 rounded-full h-6 flex overflow-hidden shadow-inner">
                          {totalCompleted > 0 && <div className="bg-emerald-500 h-6 flex items-center justify-center text-xs text-white font-bold transition-all" style={{ width: `${completionRate}%` }}>{completionRate > 10 ? `${englishToArabic(completionRate)}%` : ''}</div>}
                          {(totalActionable - totalCompleted) > 0 && <div className="bg-red-400 h-6 flex items-center justify-center text-xs text-white font-bold transition-all" style={{ width: `${100 - completionRate}%` }}>{(100 - completionRate) > 10 ? `${englishToArabic(100 - completionRate)}%` : ''}</div>}
                      </div>
                      <div className="mt-4 text-center bg-slate-50 p-2 rounded-xl border border-slate-100">
                          <p className="text-xs text-slate-500 font-bold">إجمالي السجلات: <span className="text-base text-slate-800 ml-1">{englishToArabic(totalActionable)}</span></p>
                      </div>
                  </div>
               </div>

               <div className="bg-white p-4 rounded-xl border shadow-sm">
                  <h3 className="text-xs font-bold mb-3 text-purple-900 bg-purple-100/60 p-2 rounded-xl text-center">السجلات حسب السنة 📅</h3>
                  <div className="h-52">
                     <ResponsiveContainer width="100%" height="100%">
                         <BarChart data={yearData} margin={{ top: 5, right: 0, left: -20, bottom: 0 }}>
                             <defs>
                               <linearGradient id="colorIndigo" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#6366f1" stopOpacity={0.9}/><stop offset="95%" stopColor="#818cf8" stopOpacity={0.6}/></linearGradient>
                               <linearGradient id="colorEmerald" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#10b981" stopOpacity={0.9}/><stop offset="95%" stopColor="#34d399" stopOpacity={0.6}/></linearGradient>
                             </defs>
                             <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                             <XAxis dataKey="name" tickFormatter={englishToArabic} tick={{fontSize: 10, fill: '#64748b'}} axisLine={false} tickLine={false} />
                             <YAxis tickFormatter={englishToArabic} tick={{fontSize: 10, fill: '#64748b'}} axisLine={false} tickLine={false} />
                             <RechartsTooltip formatter={(value: any) => [englishToArabic(value), 'العدد']} labelFormatter={englishToArabic} cursor={{fill: '#f8fafc'}} contentStyle={{borderRadius: '12px', border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)'}} />
                             <Bar dataKey="count" radius={[6, 6, 0, 0]} barSize={24}>
                                 {yearData.map((entry, index) => (
                                     <Cell key={`cell-${index}`} fill={index % 2 === 0 ? "url(#colorIndigo)" : "url(#colorEmerald)"} />
                                 ))}
                             </Bar>
                         </BarChart>
                     </ResponsiveContainer>
                  </div>
               </div>

               <div className="bg-white p-4 rounded-xl border shadow-sm">
                  <h3 className="text-xs font-bold mb-3 text-purple-900 bg-purple-100/60 p-2 rounded-xl text-center">توزيع طلبات الشراء بالأقسام 🏢</h3>
                  <div className="h-52">
                     <ResponsiveContainer width="100%" height="100%">
                         <BarChart data={deptData} margin={{ top: 5, right: 0, left: -20, bottom: 0 }}>
                             <defs>
                               <linearGradient id="colorIndigo" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#6366f1" stopOpacity={0.9}/><stop offset="95%" stopColor="#818cf8" stopOpacity={0.6}/></linearGradient>
                               <linearGradient id="colorEmerald" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#10b981" stopOpacity={0.9}/><stop offset="95%" stopColor="#34d399" stopOpacity={0.6}/></linearGradient>
                               <linearGradient id="colorAmber" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f59e0b" stopOpacity={0.9}/><stop offset="95%" stopColor="#fbbf24" stopOpacity={0.6}/></linearGradient>
                               <linearGradient id="colorRose" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f43f5e" stopOpacity={0.9}/><stop offset="95%" stopColor="#fb7185" stopOpacity={0.6}/></linearGradient>
                               <linearGradient id="colorPurple" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#a855f7" stopOpacity={0.9}/><stop offset="95%" stopColor="#c084fc" stopOpacity={0.6}/></linearGradient>
                               <linearGradient id="colorTeal" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#14b8a6" stopOpacity={0.9}/><stop offset="95%" stopColor="#2dd4bf" stopOpacity={0.6}/></linearGradient>
                               <linearGradient id="colorBlue" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#3b82f6" stopOpacity={0.9}/><stop offset="95%" stopColor="#60a5fa" stopOpacity={0.6}/></linearGradient>
                             </defs>
                             <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                             <XAxis dataKey="name" tickFormatter={englishToArabic} tick={{fontSize: 9, fill: '#64748b'}} axisLine={false} tickLine={false} />
                             <YAxis tickFormatter={englishToArabic} tick={{fontSize: 10, fill: '#64748b'}} axisLine={false} tickLine={false} />
                             <RechartsTooltip formatter={(value: any) => [englishToArabic(value), 'العدد']} labelFormatter={englishToArabic} cursor={{fill: '#f8fafc'}} contentStyle={{borderRadius: '12px', border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)'}} />
                             <Bar dataKey="count" radius={[6, 6, 0, 0]} barSize={20}>
                                 {deptData.map((entry, index) => {
                                     const colors = ["url(#colorIndigo)", "url(#colorEmerald)", "url(#colorAmber)", "url(#colorRose)", "url(#colorPurple)", "url(#colorTeal)", "url(#colorBlue)"];
                                     return <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />;
                                 })}
                             </Bar>
                         </BarChart>
                     </ResponsiveContainer>
                  </div>
               </div>
           </div>

           <BudgetAssetsInsights budgetRows={budgetRows} assetsNewRows={assetsNewRows} />

           <AdminReportsInsights adminReports={adminReports} />

           <div className="bg-white p-4 rounded-xl border shadow-sm">
               <h3 className="text-base font-black mb-3 text-purple-900 bg-purple-100/60 p-3 rounded-xl flex items-center justify-end gap-2">
                   أحدث طلبات الشراء <Clock size={18} className="text-purple-600" />
               </h3>
               <div className="overflow-x-auto border-2 border-slate-200 rounded-xl">
                   <table className="w-full text-sm text-right font-semibold" style={{ borderCollapse: "collapse" }}>
                       <thead style={{ backgroundColor: "#4f46e5", color: "white" }}>
                           <tr>
                               <th className="py-2 px-3 font-black text-center whitespace-nowrap border-2 border-indigo-400 text-sm">الحالة</th>
                               <th className="py-2 px-3 font-black text-right whitespace-nowrap border-2 border-indigo-400 text-sm">القسم الطالب</th>
                               <th className="py-2 px-3 font-black text-center whitespace-nowrap border-2 border-indigo-400 text-sm">الكمية المنفذة</th>
                               <th className="py-2 px-3 font-black text-center whitespace-nowrap border-2 border-indigo-400 text-sm">الكمية المطلوبة</th>
                               <th className="py-2 px-3 font-black text-right whitespace-nowrap border-2 border-indigo-400 text-sm">البند</th>
                               <th className="py-2 px-3 font-black text-right whitespace-nowrap border-2 border-indigo-400 text-sm">رقم الطلب</th>
                           </tr>
                       </thead>
                       <tbody>
                           {latestPurchases.length === 0 ? (
                               <tr><td colSpan={6} className="text-center py-4 text-slate-400 border-2 border-slate-200 text-sm">لا توجد طلبات حديثة</td></tr>
                           ) : latestPurchases.map((p, i) => (
                               <tr key={i} className="hover:bg-slate-50 transition-colors">
                                   <td className="py-2 px-3 text-center whitespace-nowrap border-2 border-slate-200">
                                       <span className={`px-3 py-1 rounded-full text-[10px] font-black border border-slate-300 ${
                                           p.executor === "تم التنفيذ" ? "bg-emerald-100 text-emerald-700" :
                                           p.executor === "ملغى" ? "bg-red-100 text-red-700" :
                                           "bg-amber-100 text-amber-700"
                                       }`}>
                                           {p.executor === "تم التنفيذ" ? "✅ تم" : p.executor === "ملغى" ? "❌ ملغى" : "⏳ معلق"}
                                       </span>
                                   </td>
                                   <td className="py-2 px-3 whitespace-nowrap text-slate-900 font-extrabold border-2 border-slate-200 text-xs">{p.requesting_department}</td>
                                   <td className="py-2 px-3 text-center font-black text-emerald-700 whitespace-nowrap border-2 border-slate-200 text-sm">{englishToArabic(p.quantity_executed || "-")}</td>
                                   <td className="py-2 px-3 text-center font-black text-slate-900 whitespace-nowrap border-2 border-slate-200 text-sm">{englishToArabic(p.quantity_requested || "-")}</td>
                                   <td className="py-2 px-3 font-black text-slate-950 whitespace-nowrap border-2 border-slate-200 text-sm">{p.item_name}</td>
                                   <td className="py-2 px-3 text-slate-700 font-mono text-xs font-black whitespace-nowrap border-2 border-slate-200">{englishToArabic(p.system_request_no || "-")}</td>
                               </tr>
                           ))}
                       </tbody>
                   </table>
               </div>
           </div>
       </div>
   );
};

const DataTableTab = ({ schemaId, supabase, currentUser, logAction, showToast, setConfirmDialog }: any) => {
  const currentSchema = schemas[schemaId];
  const isViewer = currentUser?.role === "viewer";
  const [records, setRecords] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editingRecord, setEditingRecord] = useState<any>(null);
  const [saving, setSaving] = useState(false);
  const [uploadingFile, setUploadingFile] = useState(false);
  const [importProgress, setImportProgress] = useState<any>(null);
  const [showImportGuide, setShowImportGuide] = useState(false);

  const [searchText, setSearchText] = useState("");
  const [departmentFilter, setDepartmentFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [yearFilter, setYearFilter] = useState("all");

  const [sortField, setSortField] = useState("request_date");
  const [sortDir, setSortDir] = useState<"asc"|"desc">("desc");
  const [sortDropdown, setSortDropdown] = useState("");
  const [selectedIds, setSelectedIds] = useState<string[]>([]);

  const [form, setForm] = useState<any>({});
  const fileInputRef = useRef<HTMLInputElement>(null);

  const fetchRecords = useCallback(async () => {
    setLoading(true);
    try {
      const data = await fetchAllPages(() => supabase.from(currentSchema.tableName).select("*"));
      setRecords(data);
    } catch (err: any) {
      console.error(err);
      showToast("خطأ في الاتصال: " + err?.message, "error");
    }
    setLoading(false);
  }, [supabase, currentSchema.tableName, showToast]);

  useEffect(() => { 
    fetchRecords(); 
    setSearchText(""); setSelectedIds([]);
  }, [schemaId, fetchRecords]);

  const filtered = useMemo(() => {
    let result = [...records];
    if (searchText) {
      const s = searchText.toLowerCase();
      result = result.filter(r => Object.values(r).some(v => v && String(v).toLowerCase().includes(s)));
    }
    if (departmentFilter !== "all") result = result.filter(r => r.requesting_department === departmentFilter);
    if (statusFilter !== "all") result = result.filter(r => r.executor === statusFilter || r.status === statusFilter);
    if (yearFilter !== "all") result = result.filter(r => String(r.year) === yearFilter);

    result.sort((a, b) => {
      let aVal = a[sortField] || "";
      let bVal = b[sortField] || "";
      if (typeof aVal === "string") aVal = aVal.toLowerCase();
      if (typeof bVal === "string") bVal = bVal.toLowerCase();
      if (aVal < bVal) return sortDir === "asc" ? -1 : 1;
      if (aVal > bVal) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
    return result;
  }, [records, searchText, departmentFilter, statusFilter, yearFilter, sortField, sortDir]);

  const openAdd = () => {
    setEditingRecord(null);
    const newForm: any = {};
    currentSchema.fields.forEach((f: any) => {
      if (f.key === "year") newForm[f.key] = new Date().getFullYear().toString();
      else if (f.key === "executor") newForm[f.key] = "لم يتم";
      else newForm[f.key] = "";
    });
    setForm(newForm);
    setShowForm(true);
  };

  const openEdit = (record: any) => {
    setEditingRecord(record);
    const newForm: any = {};
    currentSchema.fields.forEach((f: any) => { newForm[f.key] = record[f.key] || ""; });
    setForm(newForm);
    setShowForm(true);
  };

  const save = async () => {
    if (isViewer) {
      return showToast("عذراً، لا تملك الصلاحية للقيام بهذا الإجراء", "error");
    }
    const requiredField = currentSchema.fields.find((f: any) => f.required);
    if (requiredField && !String(form[requiredField.key] || "").trim()) {
      return showToast(`حقل "${requiredField.label}" مطلوب`, "error");
    }
    setSaving(true);
    try {
      const dataToSave: any = { ...form };

      // Remove readonly fields from insert
      currentSchema.fields.forEach((f: any) => {
        if (f.readonly && !editingRecord) {
          delete dataToSave[f.key];
        }
      });

      if (dataToSave.quantity_requested && dataToSave.quantity_executed) {
        dataToSave.quantity_remaining = Number(dataToSave.quantity_requested) - Number(dataToSave.quantity_executed);
      }
      dataToSave.updated_at = new Date().toISOString();

      if (currentSchema.tableName === "admin_affairs_managers" && !editingRecord) {
        // Manager creation goes through a server function (owner-only, uses Supabase Auth admin).
        await createManagerAccount({
          data: {
            email: String(dataToSave.email || "").trim(),
            password: String(dataToSave.password || ""),
            name: String(dataToSave.name || ""),
            permissions: String(dataToSave.permissions || "مشاهد فقط"),
          },
        });
        await logAction("create", currentSchema.tableName, null);
      } else if (editingRecord) {
        const updatePayload: any = { ...dataToSave };
        // Never send password field to DB — Supabase Auth owns credentials.
        delete updatePayload.password;
        const { error } = await supabase.from(currentSchema.tableName).update(updatePayload).eq("id", editingRecord.id);
        if (error) throw error;
        await logAction("update", currentSchema.tableName, editingRecord.id);
      } else {
        const insertData: any = {
          ...dataToSave,
          created_by: currentUser?.id || null,
          created_at: new Date().toISOString(),
        };
        delete insertData.password;
        if (!insertData.id) delete insertData.id;

        const { data: insertedData, error } = await supabase.from(currentSchema.tableName).insert([insertData]).select();
        if (error) {
          console.error("Insert error:", error);
          throw error;
        }
        await logAction("create", currentSchema.tableName, insertedData?.[0]?.id || null);
      }
      setShowForm(false);
      await fetchRecords();
      showToast(editingRecord ? "تم التعديل بنجاح" : "تمت الإضافة بنجاح", "success");
    } catch (err: any) { 
      console.error("Save error:", err);
      showToast("خطأ: " + (err?.message || err?.error_description || "فشل في حفظ البيانات"), "error"); 
    }
    setSaving(false);
  };

  const deleteRecord = async (id: string) => {
    if (isViewer) {
      return showToast("عذراً، لا تملك الصلاحية للقيام بهذا الإجراء", "error");
    }
    setConfirmDialog({
      title: "تأكيد الحذف",
      msg: "هل تريد بالتأكيد حذف هذا السجل؟",
      onConfirm: async () => {
        try {
          const { error } = await supabase.from(currentSchema.tableName).delete().eq("id", id);
          if (error) throw error;
          await logAction("delete", currentSchema.tableName, id);
          await fetchRecords();
          showToast("تم الحذف بنجاح", "success");
        } catch (err: any) { 
          console.error("Delete error:", err);
          showToast("خطأ: " + err?.message, "error"); 
        }
      }
    });
  };

  const deleteSelected = async () => {
    if (isViewer) {
      return showToast("عذراً، لا تملك الصلاحية للقيام بهذا الإجراء", "error");
    }
    if (selectedIds.length === 0) return;
    setConfirmDialog({
      title: "حذف متعدد",
      msg: `⚠️ هل أنت متأكد من حذف ${englishToArabic(selectedIds.length)} سجل نهائياً؟`,
      onConfirm: async () => {
        try {
          for (const id of selectedIds) {
            const { error } = await supabase.from(currentSchema.tableName).delete().eq("id", id);
            if (error) console.error("Delete batch error:", error);
          }
          await logAction("bulk_delete", currentSchema.tableName, null);
          setSelectedIds([]);
          await fetchRecords();
          showToast("تم حذف السجلات المحددة بنجاح", "success");
        } catch (err: any) { 
          console.error("Bulk delete error:", err);
          showToast("خطأ: " + err?.message, "error"); 
        }
      }
    });
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadingFile(true);
    setImportProgress({ total: 0, done: 0, errors: [], success: false });

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { cellDates: true, cellFormula: false, cellNF: false });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];

      const schemaKeys = Object.keys(currentSchema.excelColumns).map(k => k.trim());
      let headerRowIndex = 0;
      const rawAll = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null }) as any[][];

      let headerFound = false;
      for (let i = 0; i < Math.min(rawAll.length, 10); i++) {
        const rowVals = rawAll[i].map((v: any) => String(v ?? "").trim());
        const matches = schemaKeys.filter(k => rowVals.includes(k)).length;
        if (matches >= 2) { headerRowIndex = i; headerFound = true; break; }
      }
      if (!headerFound) {
        setImportProgress({ total: 0, done: 0, errors: [{ msg: `لم يتم العثور على رأس الجدول. تأكد من أن الملف يحتوي على الأعمدة: ${schemaKeys.slice(0,4).join(" / ")}...` }], success: false });
        setUploadingFile(false);
        if (fileInputRef.current) fileInputRef.current.value = "";
        return;
      }

      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "", range: headerRowIndex }) as any[];
      const normKey = (k: string) => String(k).trim().replace(/\s+/g, " ");
      const colMap: Record<string, string> = {};
      Object.entries(currentSchema.excelColumns).forEach(([exCol, dbField]: any) => { colMap[normKey(exCol)] = dbField; });

      const dateFields = currentSchema.fields.filter((f: any) => f.type === "date").map((f: any) => f.key);
      const toInsert: any[] = [];
      const rowErrors: any[] = [];

      jsonData.forEach((row: any, idx: number) => {
        const record: any = {};
        const cleanRow: Record<string, any> = {};
        Object.entries(row).forEach(([k, v]) => { cleanRow[normKey(k)] = v; });

        Object.entries(colMap).forEach(([exNorm, dbField]) => {
          let value = cleanRow[exNorm];
          if (value === undefined || value === null) value = "";
          if (typeof value === "string" && value.startsWith("=")) value = "";

          if (dateFields.includes(dbField)) {
            const parsed = parseImportDate(value);
            record[dbField] = parsed ?? null;
          } else if (dbField !== "quantity_remaining") {
            if (["quantity_requested","quantity_executed","year"].includes(dbField)) {
              const n = parseFloat(arabicToEnglish(String(value)));
              record[dbField] = isNaN(n) ? null : n;
            } else {
              record[dbField] = value === "" ? null : value;
            }
          }
        });

        const req = Number(record.quantity_requested);
        const exe = Number(record.quantity_executed ?? 0);
        if (!isNaN(req)) record.quantity_remaining = req - exe;
        record.created_by = currentUser?.id || null;
        record.created_at = new Date().toISOString();
        record.updated_at = new Date().toISOString();
        // Remove id to let Supabase generate
        delete record.id;

        const requiredField = currentSchema.fields.find((f: any) => f.required);
        if (requiredField && !String(record[requiredField.key] || "").trim()) {
           rowErrors.push({ row: idx + headerRowIndex + 2, msg: `الحقل "${requiredField.label}" فارغ` });
           return;
        }
        toInsert.push(record);
      });

      if (toInsert.length === 0) {
        setImportProgress({ total: 0, done: 0, errors: [{msg: "لا توجد بيانات صالحة"}], success: false });
        setUploadingFile(false); return;
      }

      const BATCH = 10;
      let inserted = 0;
      let updated = 0;
      const uniqueKeys: string[] = currentSchema.uniqueKey || [];

      for (let i = 0; i < toInsert.length; i += BATCH) {
        const batch = toInsert.slice(i, i + BATCH);
        if (uniqueKeys.length > 0) {
          for (const rec of batch) {
            const existing = records.find((r: any) => uniqueKeys.every(k => r[k] != null && rec[k] != null && String(r[k]).trim() === String(rec[k]).trim()));
            if (existing) {
              const { error } = await supabase.from(currentSchema.tableName).update({ ...rec, updated_at: new Date().toISOString() }).eq("id", existing.id);
              if (!error) updated++;
              else console.error("Update error:", error);
            } else {
              const { error } = await supabase.from(currentSchema.tableName).insert([rec]);
              if (!error) inserted++;
              else console.error("Insert error:", error);
            }
          }
        } else {
          const { error } = await supabase.from(currentSchema.tableName).insert(batch);
          if (!error) inserted += batch.length;
          else {
            console.error("Batch insert error:", error);
            rowErrors.push({ msg: error.message });
          }
        }
        setImportProgress({ total: toInsert.length, done: inserted + updated, errors: rowErrors, success: false });
      }
      setImportProgress({ total: toInsert.length, done: inserted + updated, errors: rowErrors, success: true, inserted, updated });
      await fetchRecords();
      await logAction("bulk_import", currentSchema.tableName, null);
    } catch (err: any) {
      console.error("Import error:", err);
      setImportProgress({ total: 0, done: 0, errors: [{msg: err?.message || "فشل في الاستيراد"}], success: false });
    }
    setUploadingFile(false);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const exportToExcel = () => {
    if (filtered.length === 0) return showToast("لا توجد بيانات للتصدير", "error");
    const exportData = filtered.map(r => {
      const row: any = {};
      Object.entries(currentSchema.excelColumns).forEach(([excelCol, dbField]: any) => {
        let value = r[dbField];
        if (dbField.includes("date") && value) value = String(value).split("T")[0];
        row[excelCol] = value ?? "";
      });
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, currentSchema.title);
    XLSX.writeFile(wb, `${currentSchema.title}_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const downloadTemplate = () => {
    const templateData: Record<string, any>[] = [{}];
    currentSchema.fields.forEach((f: any) => {
      templateData[0][Object.keys(currentSchema.excelColumns).find(key => currentSchema.excelColumns[key] === f.key) || f.key] = "";
    });
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "نموذج");
    XLSX.writeFile(wb, `نموذج_${currentSchema.title}.xlsx`);
  };

  return (
    <div className="space-y-4">
      <div style={{ display: "flex", gap: "6px", flexWrap: "wrap", marginBottom: "16px" }}>
        {!isViewer && (
          <button onClick={openAdd} style={{ display: "flex", alignItems: "center", gap: "5px", background: "#4f46e5", color: "white", border: "none", borderRadius: "8px", padding: "8px 12px", fontWeight: "800", cursor: "pointer", fontSize: "13px" }}>
            <Plus size={14} /> إضافة جديد
          </button>
        )}
        {schemaId !== "users" && (
          <>
            {!isViewer && (
              <>
                <button onClick={downloadTemplate} style={{ display: "flex", alignItems: "center", gap: "5px", background: "#f59e0b", color: "white", border: "none", borderRadius: "8px", padding: "8px 12px", fontWeight: "800", cursor: "pointer", fontSize: "13px" }}>
                  <FileDown size={14} /> نموذج Excel
                </button>
                <button onClick={() => fileInputRef.current?.click()} disabled={uploadingFile} style={{ display: "flex", alignItems: "center", gap: "5px", background: uploadingFile ? "#94a3b8" : "#0284c7", color: "white", border: "none", borderRadius: "8px", padding: "8px 12px", fontWeight: "800", cursor: uploadingFile ? "not-allowed" : "pointer", fontSize: "13px" }}>
                  <Upload size={14} /> {uploadingFile ? "جاري..." : "استيراد"}
                </button>
              </>
            )}
            <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: "5px", background: "#10b981", color: "white", border: "none", borderRadius: "8px", padding: "8px 12px", fontWeight: "800", cursor: "pointer", fontSize: "13px" }}>
              <Download size={14} /> تصدير
            </button>
            <button onClick={() => setShowImportGuide(true)} style={{ display: "flex", alignItems: "center", gap: "5px", background: "#f1f5f9", color: "#475569", border: "1px solid #e2e8f0", borderRadius: "8px", padding: "8px 12px", fontWeight: "800", cursor: "pointer", fontSize: "13px" }}>
              ❓ تعليمات
            </button>
          </>
        )}
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{ display: "none" }} />
      </div>

      {importProgress && (
        <div style={{ background: "white", borderRadius: "12px", padding: "16px", border: `2px solid ${importProgress.success ? "#22c55e" : importProgress.errors.length > 0 ? "#f59e0b" : "#3b82f6"}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "10px" }}>
            <p style={{ margin: 0, fontWeight: "900", fontSize: "14px" }}>{importProgress.success ? "✅ اكتمل الاستيراد" : "📊 نتيجة الاستيراد"}</p>
            <button onClick={() => setImportProgress(null)} style={{ background: "none", border: "none", cursor: "pointer" }}>✕</button>
          </div>
          <div style={{ marginBottom: "10px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: "11px", color: "#64748b", fontWeight: "700" }}>
              <span>تم رفع {importProgress.done} من {importProgress.total}</span>
              <span>{importProgress.total ? Math.round((importProgress.done / importProgress.total) * 100) : 0}%</span>
            </div>
            <div style={{ background: "#e2e8f0", borderRadius: "99px", height: "8px", overflow: "hidden" }}>
              <div style={{ height: "100%", width: `${importProgress.total ? (importProgress.done / importProgress.total) * 100 : 0}%`, background: importProgress.success ? "#22c55e" : "#4f46e5" }} />
            </div>
          </div>
          {importProgress.errors.length > 0 && (
            <div style={{ background: "#fef9ec", borderRadius: "8px", padding: "10px", maxHeight: "140px", overflowY: "auto", fontSize: "11px", color: "#78350f" }}>
              {importProgress.errors.map((e: any, i: number) => <div key={i}>{e.msg}</div>)}
            </div>
          )}
        </div>
      )}

      <div style={{ background: "white", borderRadius: "10px", padding: "10px 12px", border: "1px solid #e2e8f0", display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center" }}>
        <div style={{ flex: 1, minWidth: "180px", position: "relative" }}>
          <Search style={{ position: "absolute", right: "8px", top: "50%", transform: "translateY(-50%)", color: "#94a3b8" }} size={14} />
          <input type="text" placeholder="بحث..." value={searchText} onChange={e => setSearchText(e.target.value)} style={{ width: "100%", padding: "8px 28px 8px 10px", border: "1px solid #e2e8f0", borderRadius: "6px", outline: "none", fontSize: "12px" }} />
        </div>
        {schemaId !== "summary" && schemaId !== "users" && (
          <>
            <select value={departmentFilter} onChange={e => setDepartmentFilter(e.target.value)} style={{ padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: "6px", fontSize: "12px", outline: "none" }}>
              <option value="all">جميع الأقسام</option>
              {departments.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
            <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)} style={{ padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: "6px", fontSize: "12px", outline: "none" }}>
              <option value="all">جميع الحالات</option>
              {statuses.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </>
        )}
        {schemaId !== "users" && (
          <select value={yearFilter} onChange={e => setYearFilter(e.target.value)} style={{ padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: "6px", fontSize: "12px", outline: "none" }}>
            <option value="all">جميع السنوات</option>
            {years.map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        )}
      </div>

      {selectedIds.length > 0 && (
        <div style={{ background: "#eef2ff", borderRadius: "8px", padding: "10px 12px", border: "1px solid #c7d2fe", display: "flex", alignItems: "center", gap: "10px" }}>
          <span style={{ fontWeight: "700", color: "#4f46e5", fontSize: "13px" }}>✓ تم تحديد {englishToArabic(selectedIds.length)} سجل</span>
          <button onClick={deleteSelected} style={{ marginLeft: "auto", padding: "6px 12px", background: "#dc2626", color: "white", border: "none", borderRadius: "6px", fontWeight: "700", fontSize: "11px", cursor: "pointer" }}>🗑️ حذف المحدد</button>
        </div>
      )}

      {loading ? <div style={{ textAlign: "center", padding: "30px" }}><Loader2 className="animate-spin text-indigo-600" size={28} /></div> : 
        <div style={{ background: "white", borderRadius: "10px", border: "1px solid #e2e8f0", height: "calc(100vh - 200px)", overflow: "auto", boxShadow: "0 2px 10px rgba(0,0,0,0.03)", display: "flex", flexDirection: "column" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:"13px", backgroundColor:"#ffffff", flex: 1 }}>
              <thead style={{ backgroundColor:"#4f46e5", color:"white", fontWeight:"800", textAlign:"center", position:"sticky", top:0, zIndex:10 }}>
                <tr>
                  {!isViewer && (
                    <th style={{ padding:"8px 6px", textAlign:"center", whiteSpace: "nowrap", border: "2px solid #4338ca", width: "36px" }}>
                      <input type="checkbox" checked={selectedIds.length === filtered.length && filtered.length > 0} onChange={e => setSelectedIds(e.target.checked ? filtered.map(r => r.id) : [])} style={{ width: "14px", height: "14px", accentColor: "#4f46e5" }} />
                    </th>
                  )}
                  {currentSchema.fields.filter((f: any) => !["requesting_department", "notes", "remarks"].includes(f.key)).map((f: any) => (
                    <SortTh key={f.key} label={f.label} field={f.key} sortField={sortField} sortDir={sortDir} sortDropdown={sortDropdown} onSort={(fi:any, di:any)=>{setSortField(fi); setSortDir(di); setSortDropdown("");}} onClear={()=>{setSortField(""); setSortDropdown("");}} onToggle={(fi:any)=>setSortDropdown(prev=>prev===fi?"":fi)} align="right" />
                  ))}
                  <th style={{ padding:"8px 6px", textAlign:"center", whiteSpace: "nowrap", border: "2px solid #4338ca", fontSize:"13px" }}>الإجراءات</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((record: any) => {
                  const isLowRemaining = schemaId === 'purchases' && Number(record.quantity_requested) > 0 && Number(record.quantity_remaining) < (Number(record.quantity_requested) * 0.1);
                  const baseBg = isLowRemaining ? "#fef2f2" : "white";
                  const hoverBg = isLowRemaining ? "#fee2e2" : "#f8fafc";

                  let alertStatus = null;
                  let alertTooltip = "";
                  if (schemaId === "purchases" && Number(record.quantity_remaining) > 0) {
                     const today = new Date();
                     if (record.receipt_date) {
                         const receipt = new Date(record.receipt_date);
                         const diffDays = Math.ceil((receipt.getTime() - today.getTime()) / (1000 * 3600 * 24));
                         if (diffDays < 0) {
                             alertStatus = "delayed";
                             alertTooltip = "تأخر التنفيذ - تجاوز تاريخ الاستلام";
                         } else if (diffDays <= 3) {
                             alertStatus = "approaching";
                             alertTooltip = "اقتراب موعد الاستلام";
                         }
                     } else if (record.request_date) {
                         const reqDate = new Date(record.request_date);
                         const diffDays = Math.ceil((today.getTime() - reqDate.getTime()) / (1000 * 3600 * 24));
                         if (diffDays > 7) {
                             alertStatus = "delayed";
                             alertTooltip = "تأخر التنفيذ - مر أكثر من 7 أيام على الطلب";
                         }
                     }
                  }

                  const handleSendAlert = (e: any) => {
                    e.stopPropagation();
                    const subject = `تنبيه: ${alertStatus === 'delayed' ? 'تأخر تنفيذ' : 'اقتراب موعد استلام'} طلب ${record.system_request_no || record.admin_request_no || record.item_name}`;
                    const body = `يرجى العلم بأن الطلب الخاص بـ ${record.item_name} ${alertStatus === 'delayed' ? 'قد تأخر تنفيذه' : 'اقترب موعد استلامه'}.%0D%0Aرقم الطلب: ${record.system_request_no || record.admin_request_no}%0D%0Aالكمية المتبقية: ${record.quantity_remaining}`;
                    window.location.href = `mailto:?subject=${subject}&body=${body}`;
                    showToast("تم فتح تطبيق البريد لإرسال التنبيه", "success");
                  };

                  return (
                  <tr key={record.id} style={{ height: "44px", background: baseBg, transition: "background 0.2s" }} onMouseEnter={e => (e.currentTarget.style.background = hoverBg)} onMouseLeave={e => (e.currentTarget.style.background = baseBg)}>
                    {!isViewer && (
                      <td style={{ padding:"8px 6px", textAlign:"center", whiteSpace: "nowrap", border: "2px solid #cbd5e1" }}>
                        <input type="checkbox" checked={selectedIds.includes(record.id)} onChange={e => setSelectedIds(e.target.checked ? [...selectedIds, record.id] : selectedIds.filter(id => id !== record.id))} style={{ width: "14px", height: "14px", accentColor: "#4f46e5" }} />
                      </td>
                    )}
                    {currentSchema.fields.filter((f: any) => !["requesting_department", "notes", "remarks"].includes(f.key)).map((f: any) => {
                      if (f.type === "date" && record[f.key]) {
                        const d = new Date(record[f.key]);
                        if (!isNaN(d.getTime())) {
                          const day = String(d.getDate()).padStart(2, "0");
                          const month = String(d.getMonth() + 1).padStart(2, "0");
                          const year = d.getFullYear();
                          return (
                            <td key={f.key} style={{ padding:"8px 6px", textAlign:"right", whiteSpace: "nowrap", color: "#334155", fontWeight: "600", border: "2px solid #cbd5e1", fontSize: "13px" }}>
                              <span style={{ display: "inline-flex", flexDirection: "row-reverse", alignItems: "center", gap: "2px", direction: "ltr" }}>
                                <span style={{ fontFamily: "Inter, sans-serif" }}>{day}</span>
                                <span style={{ fontFamily: "Inter, sans-serif" }}>-</span>
                                <span style={{ fontFamily: "Inter, sans-serif" }}>{month}</span>
                                <span style={{ fontFamily: "Inter, sans-serif" }}>-</span>
                                <span style={{ fontFamily: "Inter, sans-serif" }}>{year}</span>
                              </span>
                            </td>
                          );
                        }
                      }
                      const rawVal = record[f.key] ?? "-";
                      return <td key={f.key} style={{ padding:"8px 6px", textAlign:"right", whiteSpace: "nowrap", color: isLowRemaining && f.key === 'quantity_remaining' ? "#dc2626" : "#334155", fontWeight: isLowRemaining && f.key === 'quantity_remaining' ? "900" : "600", border: "2px solid #cbd5e1", fontSize: "13px" }}>{rawVal === "" ? "-" : englishToArabic(rawVal)}</td>;
                    })}
                    <td style={{ padding:"8px 6px", textAlign:"center", whiteSpace: "nowrap", border: "2px solid #cbd5e1" }}>
                      <div style={{ display: "flex", gap: "6px", justifyContent: "center" }}>
                        {alertStatus && (
                          <button onClick={handleSendAlert} title={alertTooltip} style={{ padding: "4px 8px", background: alertStatus === "delayed" ? "#fef2f2" : "#fffbeb", color: alertStatus === "delayed" ? "#dc2626" : "#d97706", border: "1px solid " + (alertStatus === "delayed" ? "#fecaca" : "#fde68a"), borderRadius: "6px", cursor: "pointer", fontSize: "12px" }}>
                            {alertStatus === "delayed" ? "⚠️" : "🔔"}
                          </button>
                        )}
                        {!isViewer && (
                          <>
                            <button onClick={() => openEdit(record)} style={{ padding: "4px 8px", background: "#eff6ff", color: "#0284c7", border: "none", borderRadius: "6px", cursor: "pointer", fontSize: "12px" }}>✏️</button>
                            <button onClick={() => deleteRecord(record.id)} style={{ padding: "4px 8px", background: "#fff1f2", color: "#dc2626", border: "none", borderRadius: "6px", cursor: "pointer", fontSize: "12px" }}>🗑️</button>
                          </>
                        )}
                      </div>
                    </td>
                  </tr>
                  );
                })}
                {filtered.length < 6 && Array.from({ length: 6 - filtered.length }).map((_, i) => (
                  <tr key={`empty-${i}`} style={{ height: "44px", background: "white" }}>
                    {!isViewer && <td style={{ padding:"8px 6px", border: "2px solid #cbd5e1" }}></td>}
                    {currentSchema.fields.filter((f: any) => !["requesting_department", "notes", "remarks"].includes(f.key)).map((f: any) => (
                      <td key={f.key} style={{ padding:"8px 6px", border: "2px solid #cbd5e1" }}></td>
                    ))}
                    <td style={{ padding:"8px 6px", border: "2px solid #cbd5e1" }}></td>
                  </tr>
                ))}
              </tbody>
          </table>
        </div>
      }

      {showForm && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", backdropFilter: "blur(6px)", display: "flex", alignItems: "center", justifyContent: "center", padding: "16px", zIndex: 9999 }} onClick={() => setShowForm(false)}>
          <div style={{ background: "white", borderRadius: "16px", width: "100%", maxWidth: "700px", padding: "20px", boxShadow: "0 32px 80px rgba(0,0,0,0.25)", maxHeight: "90vh", overflowY: "auto" }} dir="rtl" onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "14px" }}>
              <h3 style={{ margin: 0, fontWeight: "900", fontSize: "16px" }}>{editingRecord ? "✏️ تعديل" : "➕ إضافة جديد"}</h3>
              <button onClick={() => setShowForm(false)} style={{ border: "1px solid #e2e8f0", borderRadius: "6px", padding: "5px 10px", cursor: "pointer", background: "white" }}>✕</button>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
              {currentSchema.fields.map((field: any) => (
                <div key={field.key} style={{ gridColumn: field.type === "textarea" ? "1 / -1" : "auto" }}>
                  <label style={{ fontSize: "11px", fontWeight: "700", color: "#64748b", display: "block", marginBottom: "4px" }}>{field.label} {field.required && "*"}</label>
                  {field.type === "select" ? (
                    <select value={form[field.key] || ""} onChange={e => setForm({ ...form, [field.key]: e.target.value })} disabled={field.readonly} style={{ width: "100%", padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: "6px", outline: "none", fontSize: "13px" }}>
                      <option value="">اختر...</option>
                      {field.options?.map((opt: string) => <option key={opt} value={opt}>{opt}</option>)}
                    </select>
                  ) : field.type === "textarea" ? (
                    <textarea value={form[field.key] || ""} onChange={e => setForm({ ...form, [field.key]: e.target.value })} disabled={field.readonly} rows={2} style={{ width: "100%", padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: "6px", outline: "none", resize: "none", fontSize: "13px" }} />
                  ) : (
                    <input type={field.type} value={form[field.key] || ""} onChange={e => setForm({ ...form, [field.key]: e.target.value })} disabled={field.readonly} style={{ width: "100%", padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: "6px", outline: "none", fontSize: "13px" }} />
                  )}
                </div>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", marginTop: "14px" }}>
              <button onClick={() => setShowForm(false)} style={{ padding: "10px", background: "#f1f5f9", color: "#64748b", border: "none", borderRadius: "6px", fontWeight: "900", cursor: "pointer", fontSize: "13px" }}>إلغاء</button>
              <button onClick={save} disabled={saving} style={{ padding: "10px", background: "#4f46e5", color: "white", border: "none", borderRadius: "6px", fontWeight: "900", cursor: "pointer", fontSize: "13px" }}>{saving ? "جاري..." : "💾 حفظ"}</button>
            </div>
          </div>
        </div>
      )}

      {showImportGuide && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: "16px" }} onClick={() => setShowImportGuide(false)}>
          <div style={{ background: "white", borderRadius: "14px", padding: "24px", maxWidth: "520px", width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
            <h3 style={{ margin: "0 0 16px", fontSize: "16px", fontWeight: "900" }}>📋 تعليمات الاستيراد</h3>
            <ul style={{ margin: 0, paddingRight: "18px", fontSize: "13px", lineHeight: "1.8", color: "#1e293b" }}>
              <li>تأكد من استخدام التواريخ بصيغة <b>YYYY-MM-DD</b></li>
              <li>حمّل النموذج أولاً لضمان مطابقة الأعمدة</li>
              <li>تجنب الخلايا التي تحتوي على معادلات</li>
            </ul>
            <button onClick={() => setShowImportGuide(false)} style={{ width: "100%", padding: "10px", background: "#4f46e5", color: "white", border: "none", borderRadius: "8px", fontWeight: "900", marginTop: "16px", cursor: "pointer" }}>فهمت</button>
          </div>
        </div>
      )}
    </div>
  );
};

// ==================== BUDGET SECTION ====================

const BUDGET_SHEETS = [
  "Maintenance 210",
  "Foods & Kitchen sup. 213",
  "Hospitalty 216",
  "Housing 214",
  "Stationary 220",
  "Uniform 211",
  "Cleaning 215",
  "Medical 219",
  "Entertainment 230",
  "Total"
];

const BUDGET_SHEET_LABELS: Record<string, string> = {
  "Maintenance 210": "صيانة 210",
  "Foods & Kitchen sup. 213": "أغذية ومطبخ 213",
  "Hospitalty 216": "ضيافة 216",
  "Housing 214": "إسكان 214",
  "Stationary 220": "قرطاسية 220",
  "Uniform 211": "يونيفورم 211",
  "Cleaning 215": "نظافة 215",
  "Medical 219": "طبي 219",
  "Entertainment 230": "ترفيه 230",
  "Total": "الإجمالي"
};

const MONTHS_AR = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];
const MONTHS_EN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

interface BudgetRow {
  id: string;
  item_code: string;
  item: string;
  unit: string;
  jan: number; feb: number; mar: number; apr: number;
  may: number; jun: number; jul: number; aug: number;
  sep: number; oct: number; nov: number; dec: number;
  total_qty: number;
  price: number;
  total_cost: number;
  sheet: string;
}

interface AssetRow {
  id: string;
  item_name: string;
  description: string;
  unit: string;
  jan: number; feb: number; mar: number; apr: number;
  may: number; jun: number; jul: number; aug: number;
  sep: number; oct: number; nov: number; dec: number;
  total_qty: number;
  department: string;
  price: number;
  total_cost: number;
}

const BudgetAiForecast = ({ sheetLabel, rows, monthlyTotals, sheetTotals }: {
  sheetLabel: string;
  rows: BudgetRow[];
  monthlyTotals: { name: string; total: number }[];
  sheetTotals: { totalCost: number; totalQty: number; count: number };
}) => {
    const [loading, setLoading] = useState(false);
  const [analysis, setAnalysis] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const scenarios = useMemo(() => {
    const base = sheetTotals.totalCost;
    return [
      { pct: 5,  cost: base * 1.05, color: "#059669", bg: "#ecfdf5", label: "متفائل جداً" },
      { pct: 10, cost: base * 1.10, color: "#10b981", bg: "#f0fdf4", label: "متفائل" },
      { pct: 15, cost: base * 1.15, color: "#3b82f6", bg: "#eff6ff", label: "معتدل" },
      { pct: 20, cost: base * 1.20, color: "#f59e0b", bg: "#fffbeb", label: "متحفظ" },
      { pct: 25, cost: base * 1.25, color: "#f97316", bg: "#fff7ed", label: "حذر" },
      { pct: 30, cost: base * 1.30, color: "#ef4444", bg: "#fef2f2", label: "الأسوأ" },
    ];
  }, [sheetTotals.totalCost]);

  const forecastMonthly = useMemo(
    () => monthlyTotals.map(m => ({
      name: m.name,
      "الحالي": Math.round(m.total * (sheetTotals.totalQty > 0 ? sheetTotals.totalCost / sheetTotals.totalQty : 0)),
      "المتوقع (+30%)": Math.round(m.total * (sheetTotals.totalQty > 0 ? sheetTotals.totalCost / sheetTotals.totalQty : 0) * 1.3),
    })),
    [monthlyTotals, sheetTotals]
  );

  const topItems = useMemo(
    () => [...rows]
      .sort((a, b) => (Number(b.total_cost) || 0) - (Number(a.total_cost) || 0))
      .slice(0, 5)
      .map(r => ({ name: r.item, cost: Number(r.total_cost) || 0, qty: Number(r.total_qty) || 0 })),
    [rows]
  );

  const runAnalysis = async () => {
    if (sheetTotals.count === 0) {
      setError("لا توجد بيانات كافية للتحليل. أضف أصنافاً أولاً.");
      return;
    }
    setLoading(true);
    setError(null);
    setAnalysis(null);

    // Simple local analysis instead of server AI
    setTimeout(() => {
      const avgCost = sheetTotals.totalCost / sheetTotals.count;
      const peakMonth = monthlyTotals.reduce((max, m) => m.total > max.total ? m : max, monthlyTotals[0]);
      const lowMonth = monthlyTotals.reduce((min, m) => m.total < min.total ? m : min, monthlyTotals[0]);

      const analysisText = `تحليل موازنة ${sheetLabel}:

` +
        `• إجمالي التكلفة الحالية: ${sheetTotals.totalCost.toLocaleString("en-US")} ج
` +
        `• عدد الأصناف: ${sheetTotals.count}
` +
        `• متوسط تكلفة الصنف: ${Math.round(avgCost).toLocaleString("en-US")} ج

` +
        `• أعلى شهر استهلاك: ${peakMonth?.name || "-"} (${peakMonth?.total || 0})
` +
        `• أقل شهر استهلاك: ${lowMonth?.name || "-"} (${lowMonth?.total || 0})

` +
        `• التوقع للعام القادم (زيادة 30%): ${(sheetTotals.totalCost * 1.3).toLocaleString("en-US")} ج
` +
        `• مقدار الزيادة المتوقعة: ${(sheetTotals.totalCost * 0.3).toLocaleString("en-US")} ج

` +
        `أعلى 5 أصناف تكلفة:
` +
        topItems.map((it, i) => `${i+1}. ${it.name}: ${it.cost.toLocaleString("en-US")} ج`).join("\n");

      setAnalysis(analysisText);
      setLoading(false);
    }, 800);
  };

  const forecast20 = sheetTotals.totalCost * 1.30;
  const increase = forecast20 - sheetTotals.totalCost;

  return (
    <div style={{
      background: "linear-gradient(135deg, #f0fdfa 0%, #ecfeff 100%)",
      borderRadius: "14px",
      padding: "14px",
      color: "#134e4a",
      boxShadow: "0 8px 24px rgba(20,184,166,0.10)",
      border: "1px solid #99f6e4",
      position: "relative",
      overflow: "hidden"
    }}>
      <div style={{
        position: "absolute", top: "-50%", right: "-20%", width: "400px", height: "400px",
        background: "radial-gradient(circle, rgba(45,212,191,0.10) 0%, transparent 70%)",
        pointerEvents: "none"
      }} />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: "10px", marginBottom: "14px", position: "relative" }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
            <div style={{
              background: "linear-gradient(135deg, #2dd4bf, #0d9488)",
              padding: "8px", borderRadius: "10px",
              boxShadow: "0 8px 20px rgba(45,212,191,0.35)"
            }}>
              <Brain size={18} color="white" />
            </div>
            <h3 style={{ margin: 0, fontSize: "15px", fontWeight: 900, color: "#134e4a" }}>
              توقع الذكاء الاصطناعي لموازنة العام القادم
            </h3>
            <span style={{
              background: "linear-gradient(135deg, #2dd4bf, #0d9488)",
              color: "white",
              padding: "2px 8px",
              borderRadius: "999px",
              fontSize: "9px",
              fontWeight: 900,
              boxShadow: "0 2px 8px rgba(13,148,136,0.35)"
            }}>AI ✨</span>
          </div>
          <p style={{ margin: 0, color: "#0f766e", fontSize: "11px", fontWeight: 700 }}>
            تحليل ذكي — {sheetLabel} — سيناريو زيادة الأسعار حتى 30%
          </p>
        </div>
        <button
          onClick={runAnalysis}
          disabled={loading}
          style={{
            display: "flex", alignItems: "center", gap: "6px",
            background: loading ? "#ccfbf1" : "linear-gradient(135deg, #2dd4bf, #0d9488)",
            color: loading ? "#0f766e" : "white",
            border: "none", borderRadius: "8px",
            padding: "6px 12px", fontWeight: 900, cursor: loading ? "not-allowed" : "pointer",
            fontSize: "11px", fontFamily: "Cairo, sans-serif",
            boxShadow: loading ? "none" : "0 4px 12px rgba(13,148,136,0.35)"
          }}
        >
          {loading ? <><Loader2 size={12} className="animate-spin" /> جاري...</> : <><Sparkles size={12} /> تحليل ذكي</>}
        </button>
      </div>

      <div style={{
        display: "grid",
        gridTemplateColumns: "1fr 1fr 1fr",
        gap: "8px",
        marginBottom: "12px",
        position: "relative"
      }}>
        <div style={{ background: "rgba(255,255,255,0.85)", padding: "8px 10px", borderRadius: "10px", border: "1px solid #99f6e4", textAlign: "center" }}>
          <p style={{ margin: 0, fontSize: "10px", fontWeight: 800, color: "#0f766e" }}>💰 التكلفة الحالية</p>
          <p style={{ margin: "4px 0 0", fontSize: "15px", fontWeight: 900, color: "#134e4a" }}>
            {sheetTotals.totalCost.toLocaleString("en-US")}
            <span style={{ fontSize: "10px", color: "#0f766e", marginRight: "4px" }}>ج</span>
          </p>
        </div>
        <div style={{ background: "linear-gradient(135deg,#ccfbf1,#99f6e4)", padding: "8px 10px", borderRadius: "10px", border: "1px solid #14b8a6", textAlign: "center" }}>
          <p style={{ margin: 0, fontSize: "10px", fontWeight: 800, color: "#115e59" }}>🔮 المتوقع (+30%)</p>
          <p style={{ margin: "4px 0 0", fontSize: "15px", fontWeight: 900, color: "#134e4a" }}>
            {forecast20.toLocaleString("en-US")}
            <span style={{ fontSize: "10px", color: "#115e59", marginRight: "4px" }}>ج</span>
          </p>
        </div>
        <div style={{ background: "linear-gradient(135deg,#e0f2fe,#bae6fd)", padding: "8px 10px", borderRadius: "10px", border: "1px solid #38bdf8", textAlign: "center" }}>
          <p style={{ margin: 0, fontSize: "10px", fontWeight: 800, color: "#075985" }}>📈 مقدار الزيادة</p>
          <p style={{ margin: "4px 0 0", fontSize: "15px", fontWeight: 900, color: "#0c4a6e" }}>
            +{increase.toLocaleString("en-US")}
            <span style={{ fontSize: "10px", color: "#075985", marginRight: "4px" }}>ج</span>
          </p>
        </div>
      </div>

      <div style={{
        background: "rgba(255,255,255,0.7)", borderRadius: "10px", padding: "10px",
        border: "1px solid #99f6e4", marginBottom: "12px", position: "relative"
      }}>
        <p style={{ margin: "0 0 8px", fontSize: "11px", fontWeight: 900, color: "#134e4a" }}>
          🎯 السيناريوهات المتوقعة
        </p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: "6px" }}>
          {scenarios.map(s => (
            <div key={s.pct} style={{
              background: "#ffffff",
              padding: "8px", borderRadius: "8px",
              border: `1px solid ${s.color}66`,
              borderTop: `3px solid ${s.color}`,
              textAlign: "center"
            }}>
              <p style={{ margin: 0, fontSize: "9px", fontWeight: 800, color: "#64748b" }}>{s.label}</p>
              <p style={{ margin: "3px 0 1px", fontSize: "14px", fontWeight: 900, color: s.color }}>+{s.pct}%</p>
              <p style={{ margin: 0, fontSize: "10px", fontWeight: 800, color: "#334155" }}>
                {s.cost.toLocaleString("en-US", { maximumFractionDigits: 0 })} ج
              </p>
            </div>
          ))}
        </div>
      </div>

      {sheetTotals.totalCost > 0 && (
        <div style={{
          background: "rgba(255,255,255,0.85)", borderRadius: "10px", padding: "10px",
          border: "1px solid #99f6e4", height: "200px", marginBottom: "12px", position: "relative"
        }}>
          <p style={{ margin: "0 0 6px", fontSize: "11px", fontWeight: 900, color: "#134e4a" }}>
            📊 مقارنة التكاليف الشهرية
          </p>
          <div style={{ height: "160px" }}>
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={forecastMonthly} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                <defs>
                  <linearGradient id="colorCurrent" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.6} />
                    <stop offset="95%" stopColor="#3b82f6" stopOpacity={0.05} />
                  </linearGradient>
                  <linearGradient id="colorForecast" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.6} />
                    <stop offset="95%" stopColor="#f59e0b" stopOpacity={0.05} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#99f6e4" vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#92400e" }} axisLine={false} tickLine={false} />
                <YAxis tick={{ fontSize: 9, fill: "#92400e" }} axisLine={false} tickLine={false} />
                <RechartsTooltip
                  contentStyle={{ background: "white", border: "1px solid #99f6e4", borderRadius: "10px", color: "#134e4a" }}
                  formatter={(v: any) => Number(v).toLocaleString("en-US") + " ج"}
                />
                <Legend wrapperStyle={{ fontSize: "10px", color: "#134e4a" }} />
                <Area type="monotone" dataKey="الحالي" stroke="#3b82f6" fill="url(#colorCurrent)" strokeWidth={2} />
                <Area type="monotone" dataKey="المتوقع (+30%)" stroke="#f59e0b" fill="url(#colorForecast)" strokeWidth={2} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {error && (
        <div style={{
          background: "#fee2e2", padding: "8px 12px", borderRadius: "8px",
          border: "1px solid #fca5a5", display: "flex", alignItems: "center", gap: "6px",
          position: "relative"
        }}>
          <AlertTriangle size={14} color="#991b1b" />
          <p style={{ margin: 0, color: "#991b1b", fontWeight: 800, fontSize: "12px" }}>{error}</p>
        </div>
      )}
      {analysis && (
        <div style={{
          background: "rgba(255,255,255,0.9)",
          padding: "12px", borderRadius: "10px", border: "1px solid #99f6e4",
          position: "relative"
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: "6px", marginBottom: "8px" }}>
            <Sparkles size={14} color="#0d9488" />
            <p style={{ margin: 0, fontWeight: 900, color: "#134e4a", fontSize: "12px" }}>تحليل الذكاء الاصطناعي</p>
          </div>
          <div style={{
            whiteSpace: "pre-wrap", color: "#334155", fontSize: "12px",
            lineHeight: 1.8, fontWeight: 600
          }}>
            {analysis}
          </div>
        </div>
      )}
      {!analysis && !error && !loading && (
        <p style={{ margin: 0, color: "#0f766e", fontSize: "11px", fontWeight: 700, textAlign: "center", position: "relative" }}>
          💡 اضغط على "تحليل ذكي" للحصول على توصيات مخصصة
        </p>
      )}
    </div>
  );
};

const BudgetSection = ({ supabase, currentUser, showToast, setConfirmDialog }: any) => {
  const isViewer = currentUser?.role === "viewer";
  const [activeSheet, setActiveSheet] = useState("Maintenance 210");
  const [allData, setAllData] = useState<Record<string, BudgetRow[]>>({});
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editingRow, setEditingRow] = useState<BudgetRow | null>(null);
  const [saving, setSaving] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [searchText, setSearchText] = useState("");
  const [importProgress, setImportProgress] = useState<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isMobile = useIsMobile();
  const headerPadTop = isMobile ? 40 : 12;
  const [form, setForm] = useState<any>({});

  const TABLE = "budget_rows";

  const fetchAll = useCallback(async () => {
    setLoading(true);
    try {
      const data = await fetchAllPages(() => supabase.from(TABLE).select("*"));
      const grouped: Record<string, BudgetRow[]> = {};
      BUDGET_SHEETS.forEach(s => { grouped[s] = []; });
      (data || []).forEach((r: BudgetRow) => {
        if (!grouped[r.sheet]) grouped[r.sheet] = [];
        grouped[r.sheet].push(r);
      });
      setAllData(grouped);
    } catch (e: any) {
      showToast("خطأ في تحميل البيانات: " + e.message, "error");
    }
    setLoading(false);
  }, [supabase, showToast]);

  useEffect(() => { fetchAll(); }, [fetchAll]);

  const currentRows = useMemo(() => {
    const rows = allData[activeSheet] || [];
    if (!searchText) return rows;
    const s = searchText.toLowerCase();
    return rows.filter(r => r.item?.toLowerCase().includes(s) || String(r.item_code)?.includes(s));
  }, [allData, activeSheet, searchText]);

  const sheetTotals = useMemo(() => {
    const rows = allData[activeSheet] || [];
    return {
      totalCost: rows.reduce((sum, r) => sum + (Number(r.total_cost) || 0), 0),
      totalQty: rows.reduce((sum, r) => sum + (Number(r.total_qty) || 0), 0),
      count: rows.length
    };
  }, [allData, activeSheet]);

  const openAdd = () => {
    const emptyForm: any = {
      item_code: "", item: "", unit: "عدد",
      jan: 0, feb: 0, mar: 0, apr: 0, may: 0, jun: 0,
      jul: 0, aug: 0, sep: 0, oct: 0, nov: 0, dec: 0,
      price: 0, sheet: activeSheet
    };
    setForm(emptyForm);
    setEditingRow(null);
    setShowForm(true);
  };

  const openEdit = (row: BudgetRow) => {
    setForm({ ...row });
    setEditingRow(row);
    setShowForm(true);
  };

  const calcFromForm = (f: any) => {
    const months = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
    const total_qty = months.reduce((s, m) => s + (Number(f[m]) || 0), 0);
    const total_cost = total_qty * (Number(f.price) || 0);
    return { ...f, total_qty, total_cost };
  };

  const save = async () => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    if (!form.item?.trim()) return showToast("اسم الصنف مطلوب", "error");
    setSaving(true);
    try {
      const data = calcFromForm({ ...form, sheet: activeSheet, updated_at: new Date().toISOString() });
      // Remove id for new inserts to let Supabase generate UUID
      if (!editingRow) delete data.id;

      if (editingRow) {
        const { error } = await supabase.from(TABLE).update(data).eq("id", editingRow.id);
        if (error) throw error;
      } else {
        const { error } = await supabase.from(TABLE).insert([{ ...data, created_at: new Date().toISOString() }]);
        if (error) throw error;
      }
      setShowForm(false);
      await fetchAll();
      showToast(editingRow ? "تم التعديل" : "تمت الإضافة", "success");
    } catch (e: any) { 
      console.error("Budget save error:", e);
      showToast("خطأ: " + e.message, "error"); 
    }
    setSaving(false);
  };

  const deleteRow = (id: string) => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    setConfirmDialog({
      title: "تأكيد الحذف",
      msg: "هل تريد حذف هذا السجل؟",
      onConfirm: async () => {
        const { error } = await supabase.from(TABLE).delete().eq("id", id);
        if (error) {
          showToast("خطأ في الحذف: " + error.message, "error");
          return;
        }
        await fetchAll();
        showToast("تم الحذف", "success");
      }
    });
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    setImportProgress({ total: 0, done: 0, errors: [], success: false });
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true });
      const toInsert: BudgetRow[] = [];

      for (const sheetName of wb.SheetNames) {
        if (!BUDGET_SHEETS.includes(sheetName)) continue;
        const ws = wb.Sheets[sheetName];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as any[][];

        let headerIdx = 0;
        for (let i = 0; i < Math.min(raw.length, 5); i++) {
          const vals = raw[i].map((v: any) => String(v ?? "").trim());
          if (vals.includes("Item code") || vals.includes("Item") || vals.includes("Jan")) {
            headerIdx = i; break;
          }
        }

        const headers = (raw[headerIdx] || []).map((h: any) => String(h ?? "").trim());
        const idxOf = (name: string) => headers.indexOf(name);

        const codeIdx = idxOf("Item code");
        const itemIdx = idxOf("Item");
        const unitIdx = idxOf("Unit");
        const janIdx = idxOf("Jan");
        const priceIdx = idxOf("Price");

        for (let i = headerIdx + 1; i < raw.length; i++) {
          const row = raw[i];
          if (!row || !row[itemIdx]) continue;
          const itemVal = String(row[itemIdx] ?? "").trim();
          if (!itemVal) continue;

          const getNum = (idx: number) => {
            if (idx < 0 || row[idx] == null) return 0;
            const n = parseFloat(arabicToEnglish(String(row[idx])));
            return isNaN(n) ? 0 : n;
          };

          const BUDGET_MONTH_COLS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
          const months_qty = BUDGET_MONTH_COLS.map(mk => getNum(idxOf(mk)));
          const total_qty = months_qty.reduce((s, v) => s + v, 0);
          const price = priceIdx >= 0 ? getNum(priceIdx) : 0;
          const total_cost = total_qty * price;

          const record: any = {
            item_code: String(row[codeIdx] ?? "").trim(),
            item: itemVal,
            unit: String(row[unitIdx] ?? "عدد").trim(),
            jan: months_qty[0], feb: months_qty[1], mar: months_qty[2],
            apr: months_qty[3], may: months_qty[4], jun: months_qty[5],
            jul: months_qty[6], aug: months_qty[7], sep: months_qty[8],
            oct: months_qty[9], nov: months_qty[10], dec: months_qty[11],
            total_qty, price, total_cost,
            sheet: sheetName,
            created_at: new Date().toISOString(),
            updated_at: new Date().toISOString()
          };
          // Don't set id - let Supabase generate
          toInsert.push(record);
        }
      }

      if (toInsert.length === 0) {
        setImportProgress({ total: 0, done: 0, errors: [{ msg: "لا بيانات صالحة" }], success: false });
        setUploading(false); return;
      }

      // ── Upsert: update if item_code+sheet exists, insert if new ──
      let inserted = 0, updated = 0, done = 0;
      const errors: any[] = [];
      // build a flat lookup from already-loaded data
      const existingFlat: any[] = Object.values(allData).flat();
      for (const rec of toInsert) {
        const existing = existingFlat.find((r: any) =>
          r.sheet === rec.sheet &&
          (rec.item_code ? r.item_code === rec.item_code : r.item === rec.item)
        );
        if (existing) {
          const { error } = await supabase.from(TABLE).update({ ...rec, updated_at: new Date().toISOString() }).eq("id", existing.id);
          if (error) errors.push({ msg: `تحديث "${rec.item}": ${error.message}` });
          else updated++;
        } else {
          const { error } = await supabase.from(TABLE).insert([rec]);
          if (error) errors.push({ msg: `إضافة "${rec.item}": ${error.message}` });
          else inserted++;
        }
        done++;
        if (done % 5 === 0 || done === toInsert.length)
          setImportProgress({ total: toInsert.length, done, errors, success: false });
      }
      setImportProgress({ total: toInsert.length, done, errors, success: errors.length === 0, inserted, updated });
      await fetchAll();
    } catch (e: any) {
      setImportProgress({ total: 0, done: 0, errors: [{ msg: e.message }], success: false });
    }
    setUploading(false);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const exportToExcel = () => {
    const rows = allData[activeSheet] || [];
    if (!rows.length) return showToast("لا بيانات للتصدير", "error");
    const exportData = rows.map(r => ({
      "Item code": r.item_code, "Item": r.item, "Unit": r.unit,
      "Jan": r.jan, "Feb": r.feb, "Mar": r.mar, "Apr": r.apr,
      "May": r.may, "Jun": r.jun, "Jul": r.jul, "Aug": r.aug,
      "Sep": r.sep, "Oct": r.oct, "Nov": r.nov, "Dec": r.dec,
      "Total": r.total_qty, "Price": r.price, "Total Cost": r.total_cost
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeSheet);
    XLSX.writeFile(wb, `موازنة_${activeSheet}_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const exportAllSheets = () => {
    const wb = XLSX.utils.book_new();
    BUDGET_SHEETS.forEach(sheetName => {
      const rows = allData[sheetName] || [];
      if (!rows.length) return;
      const exportData = rows.map(r => ({
        "Item code": r.item_code, "Item": r.item, "Unit": r.unit,
        "Jan": r.jan, "Feb": r.feb, "Mar": r.mar, "Apr": r.apr,
        "May": r.may, "Jun": r.jun, "Jul": r.jul, "Aug": r.aug,
        "Sep": r.sep, "Oct": r.oct, "Nov": r.nov, "Dec": r.dec,
        "Total": r.total_qty, "Price": r.price, "Total Cost": r.total_cost
      }));
      const ws = XLSX.utils.json_to_sheet(exportData);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, `الموازنة_الكاملة_${new Date().toISOString().split("T")[0]}.xlsx`);
    showToast("تم تصدير جميع الشيتات", "success");
  };

  const deleteAllSheet = () => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    setConfirmDialog({
      title: `حذف بيانات ${BUDGET_SHEET_LABELS[activeSheet]}`,
      msg: `هل تريد حذف جميع صفوف شيت "${BUDGET_SHEET_LABELS[activeSheet]}"؟`,
      requirePassword: true,
      onConfirm: async () => {
        const rows = allData[activeSheet] || [];
        for (const r of rows) {
          await supabase.from(TABLE).delete().eq("id", r.id);
        }
        await fetchAll();
        showToast("تم حذف بيانات الشيت", "success");
      }
    });
  };

  const monthlyTotals = MONTHS_EN.map((m, i) => ({
    name: MONTHS_EN[i],
    total: (allData[activeSheet] || []).reduce((s, r) => s + (Number((r as any)[m.toLowerCase()]) || 0), 0)
  }));

  return (
    <div className="space-y-3" dir="rtl">
      <div style={{ position:"sticky", top: -headerPadTop, zIndex:50, background:"rgba(248,250,252,0.98)", backdropFilter:"blur(12px)", borderBottom:"1px solid rgba(226,232,240,0.9)", boxShadow:"0 10px 30px -12px rgba(0,0,0,0.08)", paddingTop: headerPadTop, paddingBottom:"4px", marginLeft:"-4px", marginRight:"-4px", paddingLeft:"4px", paddingRight:"4px" }}>
        <div className="bg-gradient-to-l from-amber-600 to-orange-700 text-white px-4 py-3 rounded-xl flex items-center justify-between flex-wrap gap-2" style={{ boxShadow:"0 10px 30px -12px rgba(194,65,12,0.45)" }}>
          <div>
            <h2 className="text-xl font-black flex items-center gap-2">💰 الموازنة التشغيلية</h2>
            <p className="text-orange-100 text-xs font-bold mt-0.5">إدارة موازنة الأقسام الشهرية</p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            {!isViewer && (
              <button onClick={() => fileInputRef.current?.click()} disabled={uploading}
                style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.2)", color:"white", border:"1px solid rgba(255,255,255,0.3)", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
                <Upload size={13} /> {uploading ? "جاري..." : "استيراد Excel"}
              </button>
            )}
            <button onClick={exportToExcel}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.2)", color:"white", border:"1px solid rgba(255,255,255,0.3)", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Download size={13} /> تصدير الشيت
            </button>
            <button onClick={exportAllSheets}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.9)", color:"#92400e", border:"none", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <FileDown size={13} /> تصدير الكل
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{ display:"none" }} />
          </div>
        </div>

        <div style={{ marginTop:"8px", display:"flex", gap:"4px", flexWrap:"wrap", background:"rgba(255,255,255,0.96)", backdropFilter:"blur(8px)", padding:"8px", borderRadius:"10px", border:"1px solid #e2e8f0", boxShadow:"0 4px 12px rgba(0,0,0,0.04)" }}>
          {BUDGET_SHEETS.map(s => (
            <button key={s} onClick={() => setActiveSheet(s)}
              style={{ padding:"5px 10px", borderRadius:"6px", fontWeight:"800", fontSize:"11px", cursor:"pointer", border:"none",
                background: activeSheet === s ? "linear-gradient(90deg,#f59e0b,#d97706)" : "#f1f5f9",
                color: activeSheet === s ? "white" : "#64748b",
                boxShadow: activeSheet === s ? "0 3px 10px rgba(245,158,11,0.3)" : "none"
              }}>
              {BUDGET_SHEET_LABELS[s]}
            </button>
          ))}
        </div>

        <div style={{ marginTop:"6px", display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"6px" }}>
          <div style={{ background:"linear-gradient(135deg,#fff7ed,#fed7aa)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #fdba74", textAlign:"center" }}>
            <p style={{ margin:0, color:"#92400e", fontWeight:"800", fontSize:"9px" }}>عدد الأصناف</p>
            <h3 style={{ margin:"2px 0 0", color:"#c2410c", fontSize:"13px", fontWeight:"900" }}>{Number(sheetTotals.count).toLocaleString("en-US")}</h3>
          </div>
          <div style={{ background:"linear-gradient(135deg,#f0fdf4,#bbf7d0)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #86efac", textAlign:"center" }}>
            <p style={{ margin:0, color:"#14532d", fontWeight:"800", fontSize:"9px" }}>إجمالي الكميات</p>
            <h3 style={{ margin:"2px 0 0", color:"#15803d", fontSize:"13px", fontWeight:"900" }}>{Number(sheetTotals.totalQty).toLocaleString("en-US")}</h3>
          </div>
          <div style={{ background:"linear-gradient(135deg,#eff6ff,#bfdbfe)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #93c5fd", textAlign:"center" }}>
            <p style={{ margin:0, color:"#1e3a8a", fontWeight:"800", fontSize:"9px" }}>إجمالي التكلفة</p>
            <h3 style={{ margin:"2px 0 0", color:"#1d4ed8", fontSize:"13px", fontWeight:"900" }}>{Number(sheetTotals.totalCost).toLocaleString("en-US")} ج</h3>
          </div>
        </div>
      </div>

      {importProgress && (
        <div style={{ background:"white", borderRadius:"10px", padding:"12px", border:`2px solid ${importProgress.success ? "#22c55e" : "#3b82f6"}` }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:"6px" }}>
            <p style={{ margin:0, fontWeight:"900", fontSize:"13px" }}>{importProgress.success ? "✅ اكتمل الاستيراد" : "📊 جاري الاستيراد..."}</p>
            <button onClick={() => setImportProgress(null)} style={{ background:"none", border:"none", cursor:"pointer" }}>✕</button>
          </div>
          <div style={{ background:"#e2e8f0", borderRadius:"99px", height:"6px" }}>
            <div style={{ height:"100%", width:`${importProgress.total ? (importProgress.done/importProgress.total)*100 : 0}%`, background: importProgress.success ? "#22c55e" : "#4f46e5", borderRadius:"99px" }} />
          </div>
          <p style={{ margin:"4px 0 0", fontSize:"11px", color:"#64748b", fontWeight:"700" }}>تم: {importProgress.done} من {importProgress.total} · {(importProgress as any).inserted > 0 ? `✅ جديد: ${(importProgress as any).inserted}` : ""} · {(importProgress as any).updated > 0 ? `♻️ محدّث: ${(importProgress as any).updated}` : ""}</p>
          {importProgress.errors?.length > 0 && (
            <div style={{ background:"#fef9ec", borderRadius:"6px", padding:"6px", marginTop:"6px", fontSize:"11px" }}>
              {importProgress.errors.map((e: any, i: number) => <div key={i}>{e.msg}</div>)}
            </div>
          )}
        </div>
      )}

      <div style={{ background:"white", padding:"12px", borderRadius:"12px", border:"1px solid #e2e8f0" }}>
        <h3 style={{ margin:"0 0 8px", fontWeight:"800", fontSize:"12px", color:"#92400e" }}>📊 الكميات الشهرية - {BUDGET_SHEET_LABELS[activeSheet]}</h3>
        <div style={{ height:"160px" }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={monthlyTotals} margin={{ top:0, right:0, left:-20, bottom:0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
              <XAxis dataKey="name" tick={{ fontSize:9, fill:"#94a3b8" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize:9, fill:"#94a3b8" }} axisLine={false} tickLine={false} />
              <RechartsTooltip
                formatter={(v: any) => [v.toLocaleString("en-US"), "الكمية"]}
                contentStyle={{ borderRadius:"8px", border:"none", boxShadow:"0 4px 20px rgba(0,0,0,0.1)" }}
              />
              <Bar dataKey="total" radius={[6,6,0,0]} barSize={16} fill="#f59e0b" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <BudgetAiForecast
        sheetLabel={BUDGET_SHEET_LABELS[activeSheet]}
        rows={allData[activeSheet] || []}
        monthlyTotals={monthlyTotals}
        sheetTotals={sheetTotals}
      />

      <div style={{ display:"flex", gap:"8px", alignItems:"center", flexWrap:"wrap" }}>
        <div style={{ flex:1, minWidth:"180px", position:"relative" }}>
          <Search size={14} style={{ position:"absolute", right:"8px", top:"50%", transform:"translateY(-50%)", color:"#94a3b8" }} />
          <input type="text" placeholder="بحث في الصنف أو الكود..." value={searchText} onChange={e => setSearchText(e.target.value)}
            style={{ width:"100%", padding:"8px 28px 8px 10px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontSize:"12px", fontFamily:"Cairo,sans-serif" }} />
        </div>
        {!isViewer && (
          <>
            <button onClick={openAdd}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"#f59e0b", color:"white", border:"none", borderRadius:"8px", padding:"8px 12px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Plus size={14} /> إضافة صنف
            </button>
            <button onClick={deleteAllSheet}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"#fee2e2", color:"#dc2626", border:"none", borderRadius:"8px", padding:"8px 12px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Trash2 size={14} /> حذف الكل
            </button>
          </>
        )}
      </div>

      {loading ? (
        <div style={{ textAlign:"center", padding:"40px" }}><Loader2 className="animate-spin text-amber-500" size={32} /></div>
      ) : (
        <div style={{ background:"white", borderRadius:"12px", border:"1px solid #e2e8f0", overflow:"auto", maxHeight:"calc(100vh - 280px)", boxShadow:"0 2px 10px rgba(0,0,0,0.03)" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:"12px", minWidth:"1000px" }}>
            <thead style={{ background:"linear-gradient(90deg,#b45309,#d97706)", color:"white", position:"sticky", top:0, zIndex:5 }}>
              <tr>
                <th style={{ padding:"5px 6px", textAlign:"right", border:"1px solid #b45309", fontWeight:"800", fontSize:"12px", whiteSpace:"nowrap" }}>الكود</th>
                <th style={{ padding:"5px 6px", textAlign:"right", border:"1px solid #b45309", fontWeight:"800", fontSize:"12px" }}>اسم الصنف</th>
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #b45309", fontWeight:"800", fontSize:"12px" }}>الوحدة</th>
                {MONTHS_AR.map(m => (
                  <th key={m} style={{ padding:"5px 3px", textAlign:"center", border:"1px solid #b45309", fontWeight:"800", fontSize:"10px", whiteSpace:"nowrap" }}>{m}</th>
                ))}
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #b45309", fontWeight:"800", fontSize:"12px", background:"rgba(0,0,0,0.15)" }}>المجموع</th>
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #b45309", fontWeight:"800", fontSize:"12px" }}>السعر</th>
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #b45309", fontWeight:"800", fontSize:"12px", background:"rgba(0,0,0,0.15)" }}>التكلفة</th>
                {!isViewer && <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #b45309", fontWeight:"800", fontSize:"12px" }}>إجراءات</th>}
              </tr>
            </thead>
            <tbody>
              {currentRows.length === 0 ? (
                <tr><td colSpan={17 + (!isViewer ? 1 : 0)} style={{ textAlign:"center", padding:"30px", color:"#94a3b8", fontWeight:"700", fontSize:"13px" }}>لا توجد بيانات - قم بالاستيراد أو الإضافة اليدوية</td></tr>
              ) : currentRows.map((row, i) => {
                const bg = i % 2 === 0 ? "white" : "#fffbeb";
                return (
                  <tr key={row.id} style={{ background: bg, transition:"background 0.2s" }}
                    onMouseEnter={e => (e.currentTarget.style.background = "#fef3c7")}
                    onMouseLeave={e => (e.currentTarget.style.background = bg)}>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", color:"#64748b", fontSize:"11px", fontFamily:"monospace" }}>{row.item_code}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", fontWeight:"700", color:"#1e293b", fontSize:"12px" }}>{row.item}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", color:"#64748b", fontSize:"11px" }}>{row.unit}</td>
                    {(["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"] as const).map(m => (
                      <td key={m} style={{ padding:"4px 3px", border:"1px solid #e2e8f0", textAlign:"center", color: Number((row as any)[m]) > 0 ? "#1d4ed8" : "#cbd5e1", fontWeight: Number((row as any)[m]) > 0 ? "700" : "400", fontSize:"12px" }}>
                        {Number((row as any)[m]) || "-"}
                      </td>
                    ))}
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", fontWeight:"900", color:"#92400e", fontSize:"13px", background:"#fffbeb" }}>{row.total_qty?.toLocaleString("en-US")}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", color:"#15803d", fontWeight:"700", fontSize:"12px" }}>{Number(row.price)?.toLocaleString("en-US")}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", fontWeight:"900", color:"#1d4ed8", fontSize:"13px", background:"#eff6ff" }}>{Number(row.total_cost)?.toLocaleString("en-US")}</td>
                    {!isViewer && (
                      <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center" }}>
                        <div style={{ display:"flex", gap:"4px", justifyContent:"center" }}>
                          <button onClick={() => openEdit(row)} style={{ padding:"4px 8px", background:"#eff6ff", color:"#0284c7", border:"none", borderRadius:"6px", cursor:"pointer", fontSize:"12px" }}>✏️</button>
                          <button onClick={() => deleteRow(row.id)} style={{ padding:"4px 8px", background:"#fff1f2", color:"#dc2626", border:"none", borderRadius:"6px", cursor:"pointer", fontSize:"12px" }}>🗑️</button>
                        </div>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {showForm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", backdropFilter:"blur(6px)", display:"flex", alignItems:"center", justifyContent:"center", padding:"16px", zIndex:9999 }} onClick={() => setShowForm(false)}>
          <div style={{ background:"white", borderRadius:"16px", width:"100%", maxWidth:"780px", padding:"20px", boxShadow:"0 32px 80px rgba(0,0,0,0.25)", maxHeight:"90vh", overflowY:"auto" }} dir="rtl" onClick={e => e.stopPropagation()}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"14px" }}>
              <h3 style={{ margin:0, fontWeight:"900", fontSize:"16px", color:"#92400e" }}>{editingRow ? "✏️ تعديل صنف" : "➕ إضافة صنف جديد"}</h3>
              <button onClick={() => setShowForm(false)} style={{ border:"1px solid #e2e8f0", borderRadius:"6px", padding:"5px 10px", cursor:"pointer" }}>✕</button>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:"10px", marginBottom:"12px" }}>
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>الكود</label>
                <input value={form.item_code || ""} onChange={e => setForm({...form, item_code: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }} />
              </div>
              <div style={{ gridColumn:"span 2" }}>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>اسم الصنف *</label>
                <input value={form.item || ""} onChange={e => setForm({...form, item: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }} />
              </div>
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>الوحدة</label>
                <select value={form.unit || "عدد"} onChange={e => setForm({...form, unit: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }}>
                  {units.map(u => <option key={u}>{u}</option>)}
                </select>
              </div>
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>السعر</label>
                <input type="number" value={form.price || 0} onChange={e => setForm({...form, price: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }} />
              </div>
            </div>
            <div style={{ background:"#fffbeb", borderRadius:"10px", padding:"12px", border:"1px solid #fed7aa" }}>
              <p style={{ margin:"0 0 8px", fontWeight:"800", fontSize:"12px", color:"#92400e" }}>الكميات الشهرية</p>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(6,1fr)", gap:"6px" }}>
                {MONTHS_AR.map((m, i) => {
                  const key = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"][i];
                  return (
                    <div key={m}>
                      <label style={{ fontSize:"10px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"2px" }}>{m}</label>
                      <input type="number" value={form[key] || 0} onChange={e => setForm({...form, [key]: Number(e.target.value)})} style={{ width:"100%", padding:"6px", border:"1px solid #e2e8f0", borderRadius:"5px", outline:"none", textAlign:"center", boxSizing:"border-box" }} />
                    </div>
                  );
                })}
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"10px", marginTop:"12px" }}>
              <button onClick={() => setShowForm(false)} style={{ padding:"10px", background:"#f1f5f9", border:"none", borderRadius:"6px", fontWeight:"900", cursor:"pointer", fontFamily:"Cairo,sans-serif" }}>إلغاء</button>
              <button onClick={save} disabled={saving} style={{ padding:"10px", background:"#f59e0b", color:"white", border:"none", borderRadius:"6px", fontWeight:"900", cursor:"pointer", fontFamily:"Cairo,sans-serif" }}>{saving ? "جاري..." : "💾 حفظ"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// ==================== ASSETS SECTION ====================

const AssetsSection = ({ supabase, currentUser, showToast, setConfirmDialog }: any) => {
  const isViewer = currentUser?.role === "viewer";
  const [records, setRecords] = useState<AssetRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editingRow, setEditingRow] = useState<AssetRow | null>(null);
  const [saving, setSaving] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [searchText, setSearchText] = useState("");
  const [activeDept, setActiveDept] = useState("all");
  const [importProgress, setImportProgress] = useState<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isMobile = useIsMobile();
  const headerPadTop = isMobile ? 40 : 12;
  const [form, setForm] = useState<any>({});

  const TABLE = "assets_rows";

  const fetchAll = useCallback(async () => {
    setLoading(true);
    try {
      const data = await fetchAllPages(() => supabase.from(TABLE).select("*"));
      setRecords(data);
    } catch (e: any) {
      showToast("خطأ في تحميل البيانات: " + e.message, "error");
    }
    setLoading(false);
  }, [supabase, showToast]);

  useEffect(() => { fetchAll(); }, [fetchAll]);

  const filtered = useMemo(() => {
    let r = [...records];
    if (searchText) {
      const s = searchText.toLowerCase();
      r = r.filter(x => x.item_name?.toLowerCase().includes(s) || x.description?.toLowerCase().includes(s) || x.department?.toLowerCase().includes(s));
    }
    if (activeDept !== "all") r = r.filter(x => x.department === activeDept);
    return r;
  }, [records, searchText, deptFilter]);

  const stats = useMemo(() => ({
    count: filtered.length,
    totalQty: filtered.reduce((s, r) => s + (Number(r.total_qty) || 0), 0),
    totalCost: filtered.reduce((s, r) => s + (Number(r.total_cost) || 0), 0),
  }), [filtered]);

  const deptData = useMemo(() => {
    const m: Record<string, number> = {};
    records.forEach(r => { m[r.department || "غير محدد"] = (m[r.department || "غير محدد"] || 0) + 1; });
    return Object.entries(m).map(([name, count]) => ({ name, count })).sort((a,b) => b.count - a.count);
  }, [records]);

  const openAdd = () => {
    setForm({ item_name:"", description:"", unit:"عدد", department:"Admin",
      jan:0, feb:0, mar:0, apr:0, may:0, jun:0, jul:0, aug:0, sep:0, oct:0, nov:0, dec:0, price:0 });
    setEditingRow(null);
    setShowForm(true);
  };

  const openEdit = (row: AssetRow) => {
    setForm({ ...row });
    setEditingRow(row);
    setShowForm(true);
  };

  const calcFromForm = (f: any) => {
    const months = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
    const total_qty = months.reduce((s, m) => s + (Number(f[m]) || 0), 0);
    const total_cost = total_qty * (Number(f.price) || 0);
    return { ...f, total_qty, total_cost };
  };

  const save = async () => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    if (!form.item_name?.trim()) return showToast("اسم الصنف مطلوب", "error");
    setSaving(true);
    try {
      const data = calcFromForm({ ...form, updated_at: new Date().toISOString() });
      // Remove id for new inserts
      if (!editingRow) delete data.id;

      if (editingRow) {
        const { error } = await supabase.from(TABLE).update(data).eq("id", editingRow.id);
        if (error) throw error;
      } else {
        const { error } = await supabase.from(TABLE).insert([{ ...data, created_at: new Date().toISOString() }]);
        if (error) throw error;
      }
      setShowForm(false);
      await fetchAll();
      showToast(editingRow ? "تم التعديل" : "تمت الإضافة", "success");
    } catch (e: any) { 
      console.error("Assets save error:", e);
      showToast("خطأ: " + e.message, "error"); 
    }
    setSaving(false);
  };

  const deleteRow = (id: string) => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    setConfirmDialog({
      title: "تأكيد الحذف",
      msg: "هل تريد حذف هذا السجل؟",
      onConfirm: async () => {
        const { error } = await supabase.from(TABLE).delete().eq("id", id);
        if (error) {
          showToast("خطأ في الحذف: " + error.message, "error");
          return;
        }
        await fetchAll();
        showToast("تم الحذف", "success");
      }
    });
  };

  const deleteAll = () => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    setConfirmDialog({
      title: "حذف جميع الأصول",
      msg: "هل تريد حذف جميع بيانات الأصول؟ هذا لا يمكن التراجع عنه!",
      requirePassword: true,
      onConfirm: async () => {
        for (const r of records) {
          await supabase.from(TABLE).delete().eq("id", r.id);
        }
        await fetchAll();
        showToast("تم حذف جميع البيانات", "success");
      }
    });
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    setImportProgress({ total: 0, done: 0, errors: [], success: false });
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as any[][];

      let headerIdx = 0;
      for (let i = 0; i < Math.min(raw.length, 5); i++) {
        const vals = raw[i].map((v: any) => String(v ?? "").trim());
        if (vals.includes("Items") || vals.includes("Jan") || vals.includes("يناير")) {
          headerIdx = i; break;
        }
      }

      const headers = (raw[headerIdx] || []).map((h: any) => String(h ?? "").trim());
      const idxOf = (name: string) => headers.indexOf(name);

      const itemIdx = idxOf("Items");
      const descIdx = idxOf("Description");
      const unitIdx = idxOf("Unit");
      const janIdx = idxOf("Jan");

      const toInsert: any[] = [];

      for (let i = headerIdx + 1; i < raw.length; i++) {
        const row = raw[i];
        if (!row) continue;
        const itemVal = String(row[itemIdx] ?? "").trim();
        if (!itemVal || itemVal === "Items") continue;

        const getNum = (idx: number) => {
          if (idx < 0 || row[idx] == null) return 0;
          const n = parseFloat(arabicToEnglish(String(row[idx])));
          return isNaN(n) ? 0 : n;
        };

        const ASSET_MONTH_COLS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
        const months_qty = ASSET_MONTH_COLS.map(mk => getNum(idxOf(mk)));
        const total_qty = months_qty.reduce((s, v) => s + v, 0);

        const deptIdx = headers.findIndex(h => ["department","dept"].includes(h.toLowerCase()));
        const priceIdx = headers.findIndex(h => h.toLowerCase() === "price");
        const totalCostIdx = headers.findIndex(h => ["total cost","total_cost","total"].includes(h.toLowerCase()));

        const dept = deptIdx >= 0 ? String(row[deptIdx] ?? "Admin").trim() : "Admin";
        const price = getNum(priceIdx);
        const total_cost = getNum(totalCostIdx) || total_qty * price;

        toInsert.push({
          item_name: itemVal,
          description: String(row[descIdx] ?? "").trim(),
          unit: String(row[unitIdx] ?? "عدد").trim(),
          jan: months_qty[0], feb: months_qty[1], mar: months_qty[2],
          apr: months_qty[3], may: months_qty[4], jun: months_qty[5],
          jul: months_qty[6], aug: months_qty[7], sep: months_qty[8],
          oct: months_qty[9], nov: months_qty[10], dec: months_qty[11],
          total_qty, department: dept || "Admin", price, total_cost,
          created_at: new Date().toISOString(),
          updated_at: new Date().toISOString()
        });
      }

      if (!toInsert.length) {
        setImportProgress({ total:0, done:0, errors:[{msg:"لا بيانات صالحة"}], success:false });
        setUploading(false); return;
      }

      // ── Upsert: update if item_name+department exists, insert if new ──
      let inserted = 0, updated = 0, done = 0;
      const errors: any[] = [];
      for (const rec of toInsert) {
        const existing = records.find((r: any) =>
          r.item_name === rec.item_name && r.department === rec.department
        );
        if (existing) {
          const { error } = await supabase.from(TABLE).update({ ...rec, updated_at: new Date().toISOString() }).eq("id", existing.id);
          if (error) errors.push({ msg: `تحديث "${rec.item_name}": ${error.message}` });
          else updated++;
        } else {
          const { error } = await supabase.from(TABLE).insert([rec]);
          if (error) errors.push({ msg: `إضافة "${rec.item_name}": ${error.message}` });
          else inserted++;
        }
        done++;
        if (done % 5 === 0 || done === toInsert.length)
          setImportProgress({ total: toInsert.length, done, errors, success: false });
      }
      setImportProgress({ total: toInsert.length, done, errors, success: errors.length === 0, inserted, updated });
      await fetchAll();
    } catch (e: any) {
      setImportProgress({ total:0, done:0, errors:[{msg:e.message}], success:false });
    }
    setUploading(false);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const exportToExcel = () => {
    if (!filtered.length) return showToast("لا بيانات للتصدير", "error");
    const exportData = filtered.map(r => ({
      "Items": r.item_name, "Description": r.description, "Unit": r.unit,
      "Jan": r.jan, "Feb": r.feb, "Mar": r.mar, "Apr": r.apr,
      "May": r.may, "Jun": r.jun, "Jul": r.jul, "Aug": r.aug,
      "Sep": r.sep, "Oct": r.oct, "Nov": r.nov, "Dec": r.dec,
      "Total": r.total_qty, "Department": r.department, "Price": r.price, "Total Cost": r.total_cost
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb_new = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb_new, ws, "الأصول");
    XLSX.writeFile(wb_new, `الأصول_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  // Dynamic dept tabs from actual data
  const uniqueDepts = useMemo(() => {
    return [...new Set(records.map((r: any) => r.department).filter(Boolean))].sort() as string[];
  }, [records]);

  const MONTH_KEYS_EN = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  const monthlyTotalsForDept = useMemo(() => {
    const deptRows = activeDept === "all" ? records : records.filter((r: any) => r.department === activeDept);
    return MONTHS_AR.map((name, idx) => ({
      name,
      total: deptRows.reduce((s: number, r: any) => s + (Number(r[MONTH_KEYS_EN[idx]]) || 0), 0)
    }));
  }, [records, activeDept]);

  return (
    <div className="space-y-3" dir="rtl">
      <div style={{ position:"sticky", top: -headerPadTop, zIndex:50, background:"rgba(248,250,252,0.98)", backdropFilter:"blur(12px)", borderBottom:"1px solid rgba(226,232,240,0.9)", boxShadow:"0 10px 30px -12px rgba(0,0,0,0.08)", paddingTop: headerPadTop, paddingBottom:"4px", marginLeft:"-4px", marginRight:"-4px", paddingLeft:"4px", paddingRight:"4px" }}>
        <div className="bg-gradient-to-l from-purple-700 to-violet-800 text-white px-4 py-3 rounded-xl flex items-center justify-between flex-wrap gap-2">
          <div>
            <h2 className="text-xl font-black flex items-center gap-2">🏭 الأصول</h2>
            <p className="text-purple-100 text-xs font-bold mt-0.5">سجل أصول مزارع لينة وتوزيعها الشهري</p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            {!isViewer && (
              <button onClick={() => fileInputRef.current?.click()} disabled={uploading}
                style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.2)", color:"white", border:"1px solid rgba(255,255,255,0.3)", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
                <Upload size={13} /> {uploading ? "جاري..." : "استيراد Excel"}
              </button>
            )}
            <button onClick={exportToExcel}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.9)", color:"#4c1d95", border:"none", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Download size={13} /> تصدير
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{ display:"none" }} />
          </div>
        </div>

        {/* ── Dept Tabs ── */}
        <div style={{ marginTop:"8px", display:"flex", gap:"4px", flexWrap:"wrap", background:"rgba(255,255,255,0.96)", backdropFilter:"blur(8px)", padding:"8px", borderRadius:"10px", border:"1px solid #e2e8f0", boxShadow:"0 4px 12px rgba(0,0,0,0.04)" }}>
          <button onClick={() => setActiveDept("all")} style={{ padding:"5px 10px", borderRadius:"6px", fontWeight:"800", fontSize:"11px", cursor:"pointer", border:"none",
            background: activeDept === "all" ? "linear-gradient(90deg,#7c3aed,#6d28d9)" : "#f1f5f9",
            color: activeDept === "all" ? "white" : "#64748b",
            boxShadow: activeDept === "all" ? "0 3px 10px rgba(124,58,237,0.3)" : "none" }}>
            جميع الجهات
            <span style={{ marginRight:"4px", background:"rgba(255,255,255,0.25)", borderRadius:"99px", padding:"1px 6px", fontSize:"10px" }}>{records.length}</span>
          </button>
          {uniqueDepts.map((dept: string) => {
            const cnt = records.filter((r: any) => r.department === dept).length;
            const active = activeDept === dept;
            return (
              <button key={dept} onClick={() => setActiveDept(dept)} style={{ padding:"5px 10px", borderRadius:"6px", fontWeight:"800", fontSize:"11px", cursor:"pointer", border:"none",
                background: active ? "linear-gradient(90deg,#7c3aed,#6d28d9)" : "#f1f5f9",
                color: active ? "white" : "#64748b",
                boxShadow: active ? "0 3px 10px rgba(124,58,237,0.3)" : "none" }}>
                {dept}
                <span style={{ marginRight:"4px", background: active ? "rgba(255,255,255,0.25)" : "#e2e8f0", borderRadius:"99px", padding:"1px 6px", fontSize:"10px", color: active ? "white" : "#64748b" }}>{cnt}</span>
              </button>
            );
          })}
        </div>

        {/* ── Stats row ── */}
        <div style={{ marginTop:"6px", display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:"6px" }}>
          <div style={{ background:"linear-gradient(135deg,#f5f3ff,#ddd6fe)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #c4b5fd", textAlign:"center" }}>
            <p style={{ margin:0, color:"#4c1d95", fontWeight:"800", fontSize:"9px" }}>عدد الأصناف</p>
            <h3 style={{ margin:"2px 0 0", color:"#6d28d9", fontSize:"13px", fontWeight:"900" }}>{Number(stats.count).toLocaleString("en-US")}</h3>
          </div>
          <div style={{ background:"linear-gradient(135deg,#f0fdf4,#bbf7d0)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #86efac", textAlign:"center" }}>
            <p style={{ margin:0, color:"#14532d", fontWeight:"800", fontSize:"9px" }}>إجمالي الكميات</p>
            <h3 style={{ margin:"2px 0 0", color:"#15803d", fontSize:"13px", fontWeight:"900" }}>{stats.totalQty.toLocaleString("en-US")}</h3>
          </div>
          <div style={{ background:"linear-gradient(135deg,#eff6ff,#bfdbfe)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #93c5fd", textAlign:"center" }}>
            <p style={{ margin:0, color:"#1e3a8a", fontWeight:"800", fontSize:"9px" }}>إجمالي التكلفة</p>
            <h3 style={{ margin:"2px 0 0", color:"#1d4ed8", fontSize:"13px", fontWeight:"900" }}>{stats.totalCost.toLocaleString("en-US")} ج</h3>
          </div>
        </div>
      </div>

      <div style={{ display:"flex", gap:"8px", alignItems:"center", flexWrap:"wrap" }}>
        <div style={{ flex:1, minWidth:"180px", position:"relative" }}>
          <Search size={14} style={{ position:"absolute", right:"8px", top:"50%", transform:"translateY(-50%)", color:"#94a3b8" }} />
          <input type="text" placeholder="بحث في الاسم أو الوصف..." value={searchText} onChange={e => setSearchText(e.target.value)}
            style={{ width:"100%", padding:"8px 28px 8px 10px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontSize:"12px", fontFamily:"Cairo,sans-serif" }} />
        </div>
        {!isViewer && (
          <>
            <button onClick={openAdd}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"#7c3aed", color:"white", border:"none", borderRadius:"8px", padding:"8px 12px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Plus size={14} /> إضافة صنف
            </button>
            <button onClick={deleteAll}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"#fee2e2", color:"#dc2626", border:"none", borderRadius:"8px", padding:"8px 12px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Trash2 size={14} /> حذف الكل
            </button>
          </>
        )}
      </div>

      {loading ? (
        <div style={{ textAlign:"center", padding:"40px" }}><Loader2 className="animate-spin text-purple-500" size={32} /></div>
      ) : (
        <div style={{ background:"white", borderRadius:"12px", border:"1px solid #e2e8f0", overflow:"auto", maxHeight:"calc(100vh - 260px)", boxShadow:"0 2px 10px rgba(0,0,0,0.03)" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:"12px", minWidth:"1100px" }}>
            <thead style={{ background:"linear-gradient(90deg,#4c1d95,#6d28d9)", color:"white", position:"sticky", top:0, zIndex:5 }}>
              <tr>
                <th style={{ padding:"5px 6px", textAlign:"right", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px" }}>اسم الصنف</th>
                <th style={{ padding:"5px 6px", textAlign:"right", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px" }}>الوصف</th>
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px" }}>الوحدة</th>
                {MONTHS_AR.map(m => (
                  <th key={m} style={{ padding:"5px 3px", textAlign:"center", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"10px", whiteSpace:"nowrap" }}>{m}</th>
                ))}
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px", background:"rgba(0,0,0,0.15)" }}>المجموع</th>
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px" }}>الجهة</th>
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px" }}>السعر</th>
                <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px", background:"rgba(0,0,0,0.15)" }}>التكلفة</th>
                {!isViewer && <th style={{ padding:"5px 6px", textAlign:"center", border:"1px solid #5b21b6", fontWeight:"800", fontSize:"12px" }}>إجراءات</th>}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={18 + (!isViewer ? 1 : 0)} style={{ textAlign:"center", padding:"30px", color:"#94a3b8", fontWeight:"700", fontSize:"13px" }}>لا توجد بيانات - قم بالاستيراد أو الإضافة اليدوية</td></tr>
              ) : filtered.map((row, i) => {
                const bg = i % 2 === 0 ? "white" : "#faf5ff";
                return (
                  <tr key={row.id} style={{ background: bg, transition:"background 0.2s" }}
                    onMouseEnter={e => (e.currentTarget.style.background = "#f5f3ff")}
                    onMouseLeave={e => (e.currentTarget.style.background = bg)}>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", fontWeight:"700", color:"#1e293b", maxWidth:"180px", fontSize:"12px" }}>{row.item_name}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", color:"#64748b", fontSize:"11px", maxWidth:"180px" }}>{row.description}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", color:"#64748b", fontSize:"11px" }}>{row.unit}</td>
                    {(["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"] as const).map(m => (
                      <td key={m} style={{ padding:"4px 3px", border:"1px solid #e2e8f0", textAlign:"center", color: Number((row as any)[m]) > 0 ? "#6d28d9" : "#cbd5e1", fontWeight: Number((row as any)[m]) > 0 ? "700" : "400", fontSize:"12px" }}>
                        {Number((row as any)[m]) || "-"}
                      </td>
                    ))}
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", fontWeight:"900", color:"#4c1d95", fontSize:"13px", background:"#f5f3ff" }}>{row.total_qty?.toLocaleString("en-US")}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center" }}>
                      <span style={{ background:"#f0fdf4", color:"#15803d", padding:"3px 8px", borderRadius:"16px", fontSize:"11px", fontWeight:"700" }}>{row.department}</span>
                    </td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", color:"#15803d", fontWeight:"700", fontSize:"12px" }}>{Number(row.price)?.toLocaleString("en-US")}</td>
                    <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center", fontWeight:"900", color:"#1d4ed8", fontSize:"13px", background:"#eff6ff" }}>{Number(row.total_cost)?.toLocaleString("en-US")}</td>
                    {!isViewer && (
                      <td style={{ padding:"4px 6px", border:"1px solid #e2e8f0", textAlign:"center" }}>
                        <div style={{ display:"flex", gap:"4px", justifyContent:"center" }}>
                          <button onClick={() => openEdit(row)} style={{ padding:"4px 8px", background:"#eff6ff", color:"#0284c7", border:"none", borderRadius:"6px", cursor:"pointer", fontSize:"12px" }}>✏️</button>
                          <button onClick={() => deleteRow(row.id)} style={{ padding:"4px 8px", background:"#fff1f2", color:"#dc2626", border:"none", borderRadius:"6px", cursor:"pointer", fontSize:"12px" }}>🗑️</button>
                        </div>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {showForm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", backdropFilter:"blur(6px)", display:"flex", alignItems:"center", justifyContent:"center", padding:"16px", zIndex:9999 }} onClick={() => setShowForm(false)}>
          <div style={{ background:"white", borderRadius:"16px", width:"100%", maxWidth:"780px", padding:"20px", boxShadow:"0 32px 80px rgba(0,0,0,0.25)", maxHeight:"90vh", overflowY:"auto" }} dir="rtl" onClick={e => e.stopPropagation()}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"14px" }}>
              <h3 style={{ margin:0, fontWeight:"900", fontSize:"16px", color:"#4c1d95" }}>{editingRow ? "✏️ تعديل أصل" : "➕ إضافة أصل جديد"}</h3>
              <button onClick={() => setShowForm(false)} style={{ border:"1px solid #e2e8f0", borderRadius:"6px", padding:"5px 10px", cursor:"pointer" }}>✕</button>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"10px", marginBottom:"12px" }}>
              <div style={{ gridColumn:"span 2" }}>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>اسم الصنف *</label>
                <input value={form.item_name || ""} onChange={e => setForm({...form, item_name: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }} />
              </div>
              <div style={{ gridColumn:"span 2" }}>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>الوصف</label>
                <input value={form.description || ""} onChange={e => setForm({...form, description: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }} />
              </div>
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>الوحدة</label>
                <select value={form.unit || "عدد"} onChange={e => setForm({...form, unit: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }}>
                  {units.map(u => <option key={u}>{u}</option>)}
                </select>
              </div>
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>الجهة</label>
                <input value={form.department || ""} onChange={e => setForm({...form, department: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }} placeholder="Admin" />
              </div>
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>السعر</label>
                <input type="number" value={form.price || 0} onChange={e => setForm({...form, price: e.target.value})} style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box" }} />
              </div>
            </div>
            <div style={{ background:"#f5f3ff", borderRadius:"10px", padding:"12px", border:"1px solid #ddd6fe" }}>
              <p style={{ margin:"0 0 8px", fontWeight:"800", fontSize:"12px", color:"#4c1d95" }}>الكميات الشهرية</p>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(6,1fr)", gap:"6px" }}>
                {MONTHS_AR.map((m, i) => {
                  const key = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"][i];
                  return (
                    <div key={m}>
                      <label style={{ fontSize:"10px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"2px" }}>{m}</label>
                      <input type="number" value={form[key] || 0} onChange={e => setForm({...form, [key]: Number(e.target.value)})} style={{ width:"100%", padding:"6px", border:"1px solid #e2e8f0", borderRadius:"5px", outline:"none", textAlign:"center", boxSizing:"border-box" }} />
                    </div>
                  );
                })}
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"10px", marginTop:"12px" }}>
              <button onClick={() => setShowForm(false)} style={{ padding:"10px", background:"#f1f5f9", border:"none", borderRadius:"6px", fontWeight:"900", cursor:"pointer", fontFamily:"Cairo,sans-serif" }}>إلغاء</button>
              <button onClick={save} disabled={saving} style={{ padding:"10px", background:"#7c3aed", color:"white", border:"none", borderRadius:"6px", fontWeight:"900", cursor:"pointer", fontFamily:"Cairo,sans-serif" }}>{saving ? "جاري..." : "💾 حفظ"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default function AdminAffairsSystem() {
  const [currentView, setCurrentView] = useState("login");
  const [currentUser, setCurrentUser] = useState<any>(null);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [loginData, setLoginData] = useState({ email: "", password: "" });
  const [rememberMe, setRememberMe] = useState(true);
  const [sidebarOpen, setSidebarOpen] = useState(() => typeof window !== "undefined" ? window.innerWidth >= 1024 : true);
  const [isMobile, setIsMobile] = useState(() => typeof window !== "undefined" ? window.innerWidth < 1024 : false);
  const restoreFileRef = useRef<HTMLInputElement>(null);
  const [isSystemProcessing, setIsSystemProcessing] = useState(false);
  const [showSystemMenu, setShowSystemMenu] = useState(false);
  const [toastMessage, setToastMessage] = useState<{msg: string, type: "success" | "error" | "info"} | null>(null);
  const [confirmDialog, setConfirmDialog] = useState<{title: string, msg: string, requirePassword?: boolean, onConfirm: ()=>void} | null>(null);

  const showToast = useCallback((msg: string, type: "success" | "error" | "info" = "info") => {
    setToastMessage({msg, type});
    setTimeout(() => setToastMessage(null), 3500);
  }, []);

  useEffect(() => {
    const handleResize = () => {
      const mobile = window.innerWidth < 1024;
      setIsMobile(mobile);
      if (mobile) setSidebarOpen(false);
    };
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    // Restore session from Supabase Auth and resolve role.
    let mounted = true;
    const resolveUser = async (session: any) => {
      if (!session?.user) {
        if (mounted) { setCurrentUser(null); setCurrentView("login"); }
        return;
      }
      const uid = session.user.id;
      const { data: roles } = await supabase.from("user_roles").select("role").eq("user_id", uid);
      const roleSet = new Set((roles || []).map((r: any) => r.role));
      let appRole: "owner" | "manager" | "viewer" = "viewer";
      if (roleSet.has("owner")) appRole = "owner";
      else if (roleSet.has("manager")) appRole = "manager";
      let name = session.user.email || "";
      let permissions: string | undefined;
      const { data: mgr } = await supabase.from("admin_affairs_managers").select("name,permissions").eq("user_id", uid).maybeSingle();
      if (mgr) { name = mgr.name || name; permissions = mgr.permissions || undefined; }
      if (!mounted) return;
      setCurrentUser({ id: uid, role: appRole, name, permissions });
      setCurrentView("app");
    };
    supabase.auth.getSession().then(({ data }) => resolveUser(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_e, session) => resolveUser(session));
    return () => { mounted = false; sub.subscription.unsubscribe(); };
  }, []);

  const logAction = async (action: string, tableName: string, recordId: any = null) => {
    try {
      await supabase.from("audit_log").insert([{
        user_id: currentUser?.id || null,
        user_name: currentUser?.name || "النظام",
        action, table_name: tableName, record_id: recordId,
      }]);
    } catch (err) {}
  };

  // ==================== FIXED: Backup ALL system tables ====================
  const handleBackupAll = async () => {
    try {
      setIsSystemProcessing(true);
      const backup: any = {};

      // Backup all schema tables
      for (const key of Object.keys(schemas)) {
        const { data, error } = await supabase.from(schemas[key].tableName).select("*");
        if (error) {
          console.error(`Backup error for ${schemas[key].tableName}:`, error);
          showToast(`خطأ في نسخ ${schemas[key].title}: ${error.message}`, "error");
        }
        backup[key] = data || [];
      }

      // Backup all other tables
      const otherTables = ["budget_rows", "assets_rows", "audit_log", "admin_reports"];
      for (const tableName of otherTables) {
        const { data, error } = await supabase.from(tableName as any).select("*");
        if (error) {
          console.error(`Backup error for ${tableName}:`, error);
        }
        backup[tableName] = data || [];
      }

      const dataStr = JSON.stringify(backup, null, 2);
      const blob = new Blob([dataStr], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `system_backup_${new Date().toISOString().split('T')[0]}.json`;
      a.click();
      URL.revokeObjectURL(url);
      await logAction("system_backup", "all_tables", null);
      showToast("تم النسخ الاحتياطي بنجاح لجميع الأقسام", "success");
    } catch (e: any) {
      showToast("خطأ في النسخ الاحتياطي: " + e.message, "error");
    } finally {
      setIsSystemProcessing(false);
    }
  };

  // ==================== FIXED: Restore ALL system tables ====================
  const handleRestoreAll = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setConfirmDialog({
      title: "استعادة النظام",
      msg: "سيتم استعادة النظام بالكامل من الملف المرفق، هل تريد المتابعة؟",
      onConfirm: () => {
        const reader = new FileReader();
        reader.onload = async (event) => {
          try {
            const data = JSON.parse(event.target?.result as string);
            setIsSystemProcessing(true);

            // Restore schema tables
            for (const key of Object.keys(data)) {
              if (schemas[key]) {
                const records = data[key];
                if (records && records.length > 0) {
                  // Remove existing IDs to let Supabase generate new ones
                  const cleanRecords = records.map((r: any) => {
                    const clean = { ...r };
                    delete clean.id; // Let Supabase generate new UUIDs
                    return clean;
                  });

                  const BATCH = 50;
                  for (let i = 0; i < cleanRecords.length; i += BATCH) {
                    const { error } = await supabase.from(schemas[key].tableName).insert(cleanRecords.slice(i, i + BATCH));
                    if (error) {
                      console.error(`Restore error for ${schemas[key].tableName}:`, error);
                      showToast(`خطأ في استعادة ${schemas[key].title}: ${error.message}`, "error");
                    }
                  }
                }
              }

              // Restore other tables
              if (["budget_rows", "assets_rows", "audit_log"].includes(key) && data[key]?.length > 0) {
                const cleanRecords = data[key].map((r: any) => {
                  const clean = { ...r };
                  delete clean.id;
                  return clean;
                });

                const BATCH = 50;
                for (let i = 0; i < cleanRecords.length; i += BATCH) {
                  const { error } = await supabase.from(key as any).insert(cleanRecords.slice(i, i + BATCH));
                  if (error) {
                    console.error(`Restore error for ${key}:`, error);
                  }
                }
              }
            }

            await logAction("system_restore", "all_tables", null);
            showToast("تمت الاستعادة بنجاح، يتم تحديث الصفحة...", "success");
            setTimeout(() => window.location.reload(), 1500);
          } catch (err: any) {
            showToast("خطأ في الاستعادة: " + err.message, "error");
            setIsSystemProcessing(false);
          }
        };
        reader.readAsText(file);
      }
    });

    if (e.target) e.target.value = "";
  };

  // ==================== FIXED: Delete ALL system data ====================
  const handleDeleteAllData = async () => {
    setConfirmDialog({
      title: "تحذير خطير: حذف جميع البيانات",
      msg: "هل أنت متأكد من حذف جميع بيانات النظام بالكامل؟ هذا الإجراء لا يمكن التراجع عنه بأي شكل!",
      requirePassword: true,
      onConfirm: async () => {
        try {
          setIsSystemProcessing(true);

          // Delete all schema tables
          for (const key of Object.keys(schemas)) {
            const { error } = await supabase.from(schemas[key].tableName).delete().neq("id", "00000000-0000-0000-0000-000000000000");
            if (error) {
              console.error(`Delete error for ${schemas[key].tableName}:`, error);
            }
          }

          // Delete all other tables
          const otherTables = ["budget_rows", "assets_rows", "audit_log", "admin_reports"];
          for (const tableName of otherTables) {
            const { error } = await supabase.from(tableName as any).delete().neq("id", "00000000-0000-0000-0000-000000000000");
            if (error) {
              console.error(`Delete error for ${tableName}:`, error);
            }
          }

          await logAction("system_wipe", "all_tables", null);
          showToast("تم حذف جميع البيانات بنجاح، يتم تحديث الصفحة...", "success");
          setTimeout(() => window.location.reload(), 1500);
        } catch (err: any) {
          showToast("خطأ في الحذف: " + err.message, "error");
          setIsSystemProcessing(false);
        }
      }
    });
  };

  const handleLogin = async () => {
    if (!loginData.email.trim() || !loginData.password) {
      showToast("يرجى إدخال اسم المستخدم وكلمة المرور", "error");
      return;
    }

    try {
      // Try to bootstrap the owner account on first-ever login attempt.
      // If an owner already exists in the DB, this is a no-op server-side.
      try {
        await bootstrapOwner({
          data: {
            email: loginData.email.trim(),
            password: loginData.password,
          },
        });
      } catch { /* ignore – owner already exists or not applicable */ }

      const { data, error } = await supabase.auth.signInWithPassword({
        email: loginData.email.trim(),
        password: loginData.password,
      });
      if (error || !data.session) {
        showToast("بيانات الدخول غير صحيحة، يرجى المحاولة مرة أخرى", "error");
        return;
      }
      showToast("تم تسجيل الدخول بنجاح", "success");
      // onAuthStateChange in the mount effect will resolve role and switch view.
    } catch (e: any) {
      console.error("Login error:", e);
      showToast("بيانات الدخول غير صحيحة، يرجى المحاولة مرة أخرى", "error");
    }
  };

  if (currentView === "login") {
    return (
      <div dir="rtl" style={{ 
        minHeight: "100vh", 
        background: "linear-gradient(180deg, #0f3d23 0%, #166534 25%, #22c55e 55%, #a7f3d0 85%, #052e16 100%)", 
        display: "flex", 
        flexDirection: "column",
        alignItems: "center", 
        justifyContent: "center", 
        padding: "20px",
        position: "relative",
        overflow: "hidden"
      }}>
        <div style={{ position: "absolute", bottom: 0, left: 0, right: 0, height: "30vh", overflow: "hidden", pointerEvents: "none", display: "flex", alignItems: "flex-end" }}>
          <svg viewBox="0 0 1440 300" style={{ width: "100%", height: "100%", position: "absolute", bottom: 0, left: 0 }} preserveAspectRatio="none">
            <path d="M0,300 L1440,300 L1440,240 Q1080,280 720,250 Q360,220 0,260 Z" fill="#032511" />
            <g transform="translate(80, 50) scale(1.6)" fill="#032511">
              <path d="M20,130 Q15,70 10,10 Q14,10 24,70 Q25,130 20,130" />
              <path d="M10,10 Q-5,-10 -25,-5 Q-5,0 10,10" />
              <path d="M10,10 Q25,-10 45,-5 Q25,0 10,10" />
              <path d="M10,10 Q-15,5 -30,25 Q-15,18 10,10" />
              <path d="M10,10 Q35,5 50,25 Q35,18 10,10" />
              <path d="M10,10 Q-10,25 -20,50 Q-8,35 10,10" />
              <path d="M10,10 Q30,25 40,50 Q28,35 10,10" />
            </g>
            <g transform="translate(260, 90) scale(1.3)" fill="#032511" opacity="0.9">
              <path d="M20,130 Q16,75 12,15 Q15,15 24,75 Q24,130 20,130" />
              <path d="M12,15 Q-2,-5 -20,0 Q-3,5 12,15" />
              <path d="M12,15 Q26,-5 44,0 Q26,5 12,15" />
              <path d="M12,15 Q-11,10 -25,30 Q-11,22 12,15" />
              <path d="M12,15 Q35,10 49,30 Q35,22 12,15" />
            </g>
            <g transform="translate(1080, 80) scale(1.4)" fill="#032511" opacity="0.95">
              <path d="M20,130 Q17,70 11,12 Q14,12 23,70 Q23,130 20,130" />
              <path d="M11,12 Q-4,-8 -22,-3 Q-4,2 11,12" />
              <path d="M11,12 Q26,-8 44,-3 Q26,2 11,12" />
              <path d="M11,12 Q-12,7 -26,27 Q-12,19 11,12" />
              <path d="M11,12 Q34,7 48,27 Q34,19 11,12" />
            </g>
            <g transform="translate(1260, 40) scale(1.7)" fill="#032511">
              <path d="M20,130 Q14,65 8,8 Q12,8 24,65 Q26,130 20,130" />
              <path d="M8,8 Q-7,-12 -27,-7 Q-7,-2 8,8" />
              <path d="M8,8 Q23,-12 43,-7 Q23,-2 8,8" />
              <path d="M8,8 Q-17,3 -32,23 Q-17,16 8,8" />
              <path d="M8,8 Q33,3 48,23 Q33,16 8,8" />
              <path d="M8,8 Q-12,23 -22,48 Q-10,33 8,8" />
              <path d="M8,8 Q28,23 38,48 Q26,33 8,8" />
            </g>
          </svg>
        </div>

        <div style={{ display: "flex", justifyContent: "center", marginBottom: "16px", zIndex: 10 }}>
          <img 
            src="/logo.png" 
            alt="لوجو مزارع لينة" 
            style={{ width: "120px", height: "auto", objectFit: "contain", userSelect: "none" }} 
            onError={(e) => { e.currentTarget.style.display = 'none'; }} 
          />
        </div>

        <div style={{ 
          background: "rgba(255, 255, 255, 0.25)", 
          backdropFilter: "blur(24px)", 
          WebkitBackdropFilter: "blur(24px)",
          padding: "36px 32px", 
          borderRadius: "28px", 
          width: "100%", 
          maxWidth: "420px", 
          boxShadow: "0 30px 60px -15px rgba(2, 40, 15, 0.35)",
          border: "1.5px solid rgba(255, 255, 255, 0.45)",
          textAlign: "center",
          zIndex: 10
        }}>
          <h2 style={{ color: "#064e3b", fontSize: "28px", fontWeight: "900", marginBottom: "24px", textAlign: "center" }}>تسجيل الدخول</h2>

          <div style={{ position: "relative", marginBottom: "14px", width: "100%" }}>
            <input 
              placeholder="اسم المستخدم" 
              value={loginData.email} 
              onChange={e => setLoginData({...loginData, email: e.target.value})} 
              style={{ 
                width: "100%", 
                background: "rgba(255, 255, 255, 0.45)", 
                border: "2px solid #166534", 
                padding: "14px 18px 14px 44px", 
                borderRadius: "9999px", 
                color: "#064e3b", 
                outline: "none", 
                boxSizing: "border-box",
                fontWeight: "700",
                fontSize: "14px",
                textAlign: "right"
              }} 
            />
            <User 
              size={18} 
              className="text-emerald-900" 
              style={{ position: "absolute", left: "16px", top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} 
            />
          </div>

          <div style={{ position: "relative", marginBottom: "16px", width: "100%" }}>
            <input 
              type="password" 
              placeholder="كلمة المرور" 
              value={loginData.password} 
              onChange={e => setLoginData({...loginData, password: e.target.value})} 
              onKeyDown={e => e.key === "Enter" && handleLogin()} 
              style={{ 
                width: "100%", 
                background: "rgba(255, 255, 255, 0.45)", 
                border: "2px solid #166534", 
                padding: "14px 18px 14px 44px", 
                borderRadius: "9999px", 
                color: "#064e3b", 
                outline: "none", 
                boxSizing: "border-box",
                fontWeight: "700",
                fontSize: "14px",
                textAlign: "right"
              }} 
            />
            <Lock 
              size={18} 
              className="text-emerald-900" 
              style={{ position: "absolute", left: "16px", top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} 
            />
          </div>

          <div style={{ display: "flex", justifyContent: "flex-start", alignItems: "center", marginBottom: "22px", padding: "0 8px", width: "100%" }}>
            <label style={{ display: "flex", alignItems: "center", gap: "8px", color: "#064e3b", fontWeight: "700", fontSize: "14px", cursor: "pointer", userSelect: "none" }}>
              <input 
                type="checkbox" 
                checked={rememberMe} 
                onChange={e => setRememberMe(e.target.checked)} 
                style={{ accentColor: "#15803d", width: "18px", height: "18px", borderRadius: "6px", cursor: "pointer" }} 
              />
              تذكرني؟
            </label>
          </div>

          <button 
            onClick={handleLogin} 
            style={{ 
              width: "100%", 
              background: "#0d522c", 
              color: "white", 
              border: "none", 
              padding: "14px", 
              borderRadius: "9999px", 
              fontWeight: "900", 
              fontSize: "16px", 
              cursor: "pointer",
              boxShadow: "0 10px 20px -3px rgba(13, 82, 44, 0.4)",
              transition: "all 0.2s ease"
            }}
            onMouseEnter={e => { e.currentTarget.style.background = "#093d20"; e.currentTarget.style.transform = "scale(1.02)"; }}
            onMouseLeave={e => { e.currentTarget.style.background = "#0d522c"; e.currentTarget.style.transform = "none"; }}
          >
            دخول
          </button>
        </div>

        {toastMessage && (
          <div className={`fixed bottom-6 right-6 p-3 rounded-xl shadow-xl flex items-center gap-2 z-[9999] text-white font-bold animate-in slide-in-from-bottom-5 ${toastMessage.type === 'success' ? 'bg-emerald-500' : toastMessage.type === 'error' ? 'bg-red-500' : 'bg-blue-500'}`}>
            {toastMessage.type === 'success' ? <CheckCircle size={18} /> : toastMessage.type === 'error' ? <X size={18} /> : <div />}
            {toastMessage.msg}
          </div>
        )}
      </div>
    );
  }

  const tabs = [
    { id: "dashboard", label: "الرئيسية", icon: LayoutDashboard },
    { id: "purchases", label: "طلبات الشراء", icon: ShoppingCart },
    { id: "summary", label: "إجمالي الطلبات", icon: ListOrdered },
    { id: "vegetables", label: "خضار أسبوعي", icon: Leaf },
    { id: "assets", label: "الأصول", icon: Package },
    { id: "budget", label: "الموازنة", icon: BarChart2 },
    { id: "assets_new", label: "الأصول الجديد", icon: Building2 },
    { id: "admin_reports", label: "التقارير الإدارية", icon: CalendarDays },
  ];

  if (currentUser?.role === "owner" || currentUser?.role === "admin") {
    tabs.push({ id: "users", label: "المستخدمين", icon: Users });
  }

  return (
    <div className="min-h-screen flex bg-slate-50 relative" dir="rtl">
        {sidebarOpen && isMobile && (
          <div
            className="fixed inset-0 bg-black/50 z-40"
            onClick={() => setSidebarOpen(false)}
          />
        )}

        <button
          onClick={() => setSidebarOpen(!sidebarOpen)}
          className={`fixed top-3 z-50 text-white p-2.5 rounded-xl shadow-lg transition-all border-none cursor-pointer flex items-center justify-center ${sidebarOpen && !isMobile ? "right-[260px]" : "right-3"}`}
          style={{ background: sidebarOpen ? "rgba(79, 70, 229, 0.9)" : "linear-gradient(135deg, #4f46e5, #7c3aed)", backdropFilter:"blur(10px)" }}
        >
          {sidebarOpen ? <X size={20} /> : <LayoutDashboard size={20} />}
        </button>

      <aside style={{ 
        width: sidebarOpen ? "240px" : "0", 
        minWidth: sidebarOpen ? "240px" : "0", 
        background: "linear-gradient(180deg, #0f172a 0%, #1e293b 100%)", 
        padding: sidebarOpen ? "20px 12px" : "0", 
        display: "flex", 
        flexDirection: "column", 
        position: isMobile ? "fixed" : "sticky", 
        top: 0, 
        bottom: 0, 
        right: 0,
        height: "100vh",
        zIndex: 60,
        transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
        overflowY: sidebarOpen ? "auto" : "hidden",
        overflowX: "hidden",
        borderLeft: sidebarOpen ? "1px solid #334155" : "none",
        boxShadow: sidebarOpen && isMobile ? "-8px 0 32px rgba(0,0,0,0.5)" : "none",
      }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "8px", marginBottom: "20px", padding: "16px 10px", opacity: sidebarOpen ? 1 : 0, transition: "opacity 0.3s", textAlign: "center", background: "linear-gradient(160deg, rgba(16,185,129,0.12) 0%, rgba(79,70,229,0.08) 100%)", borderRadius: "16px", border: "1px solid rgba(255,255,255,0.08)", position: "relative", overflow: "hidden" }}>
          <div style={{ position: "absolute", top: "-40px", left: "-40px", width: "120px", height: "120px", background: "radial-gradient(circle, rgba(52,211,153,0.25), transparent 70%)", pointerEvents: "none" }} />
          <img src="/logo.png" alt="مشترياتك" style={{ width: "60px", height: "auto", objectFit: "contain", filter: "drop-shadow(0 6px 12px rgba(0,0,0,0.4))", position: "relative", zIndex: 1 }} onError={(e) => { e.currentTarget.style.display = 'none'; }} />
          <div style={{ position: "relative", zIndex: 1 }}>
            <h1 style={{ fontWeight: "900", fontSize: "17px", margin: "0 0 2px 0", color: "#f8fafc", letterSpacing: "0.5px" }}>مشترياتك</h1>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "4px" }}>
              <span style={{ width: "3px", height: "3px", borderRadius: "50%", background: "#34d399", boxShadow: "0 0 8px #34d399" }} />
              <p style={{ color: "#34d399", fontSize: "9px", fontWeight: "800", margin: 0, textTransform: "uppercase", letterSpacing: "1.5px" }}>Linah Farms</p>
              <span style={{ width: "3px", height: "3px", borderRadius: "50%", background: "#34d399", boxShadow: "0 0 8px #34d399" }} />
            </div>
            <p style={{ margin: "4px 0 0", fontSize: "9px", color: "#94a3b8", fontWeight: "600", letterSpacing: "0.3px" }}>نظام الشؤون الإدارية المتكامل</p>
          </div>
        </div>

        <nav style={{ flex: 1, display: "flex", flexDirection: "column", gap: "6px", opacity: sidebarOpen ? 1 : 0, transition: "opacity 0.3s" }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => { setActiveTab(t.id); if(isMobile) setSidebarOpen(false); }} style={{ display: "flex", alignItems: "center", gap: "10px", padding: "10px 12px", borderRadius: "10px", background: activeTab === t.id ? "linear-gradient(90deg, #4f46e5, #4338ca)" : "rgba(255,255,255,0.03)", color: activeTab === t.id ? "#ffffff" : "#94a3b8", border: activeTab === t.id ? "1px solid #6366f1" : "1px solid rgba(255,255,255,0.05)", fontWeight: activeTab === t.id ? "800" : "600", cursor: "pointer", fontSize: "13px", textAlign: "right", transition: "all 0.2s cubic-bezier(0.4, 0, 0.2, 1)", boxShadow: activeTab === t.id ? "0 4px 12px rgba(79, 70, 229, 0.3)" : "none" }} onMouseEnter={e => { if(activeTab !== t.id) e.currentTarget.style.background = "rgba(255,255,255,0.08)"; }} onMouseLeave={e => { if(activeTab !== t.id) e.currentTarget.style.background = "rgba(255,255,255,0.03)"; }}>
              <t.icon size={18} /> <span style={{ flex: 1 }}>{t.label}</span>
              {activeTab === t.id && <div style={{ width: "5px", height: "5px", borderRadius: "50%", background: "#fff", boxShadow: "0 0 8px #fff" }} />}
            </button>
          ))}
        </nav>

        <div style={{ marginTop: "auto", opacity: sidebarOpen ? 1 : 0, transition: "opacity 0.3s" }}>
          <div style={{ padding: "12px", background: "rgba(255,255,255,0.05)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: "14px", marginBottom: "10px", textAlign: "center", boxShadow: "inset 0 2px 10px rgba(0,0,0,0.2)" }}>
            <p style={{ margin: 0, fontWeight: "bold", color: "#f8fafc", fontSize: "13px" }}>{currentUser?.name}</p>
            <p style={{ margin: "2px 0 0 0", color: "#94a3b8", fontSize: "11px" }}>{currentUser?.department}</p>
          </div>
          <button onClick={() => { supabase.auth.signOut(); setCurrentUser(null); setCurrentView("login"); }} style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: "10px", padding: "10px 12px", borderRadius: "10px", background: "rgba(239, 68, 68, 0.1)", color: "#fca5a5", border: "1px solid rgba(239, 68, 68, 0.2)", fontWeight: "700", cursor: "pointer", fontSize: "13px", transition: "all 0.2s" }} onMouseEnter={e => { e.currentTarget.style.background = "rgba(239, 68, 68, 0.2)"; e.currentTarget.style.color = "#fff"; }} onMouseLeave={e => { e.currentTarget.style.background = "rgba(239, 68, 68, 0.1)"; e.currentTarget.style.color = "#fca5a5"; }}>
            <LogOut size={16} /> تسجيل خروج
          </button>
        </div>
      </aside>

      <main style={{ 
        flex: 1, 
        padding: isMobile ? "48px 12px 16px" : "24px 32px", 
        boxSizing: "border-box", 
        overflowY: "auto", 
        height: "100vh",
        width: "100%",
        transition: "padding 0.3s"
      }}>
        {activeTab === "dashboard" ? (
          <Dashboard 
            supabase={supabase} 
            systemMenu={
              (currentUser?.role === "admin" || currentUser?.role === "owner") ? (
                <div className="relative">
                  <input type="file" ref={restoreFileRef} style={{ display: 'none' }} accept=".json" onChange={handleRestoreAll} />
                  <button 
                    onClick={() => setShowSystemMenu(!showSystemMenu)} 
                    className="flex items-center justify-center w-12 h-12 bg-emerald-500 hover:bg-emerald-600 text-white rounded-full shadow-[0_8px_30px_rgb(0,0,0,0.12)] transition-all z-50 disabled:opacity-50 border-4 border-white/20"
                    disabled={isSystemProcessing}
                  >
                    {isSystemProcessing ? <Loader2 size={24} className="animate-spin" /> : <Save size={24} />}
                  </button>

                  {showSystemMenu && (
                    <div className="absolute top-16 left-0 w-60 bg-white rounded-2xl shadow-xl border border-slate-100 overflow-hidden z-[9999] flex flex-col gap-1 p-2">
                      <div className="p-2 mb-1 border-b border-slate-100 flex items-center gap-2 text-slate-800">
                        <Database size={14} className="text-indigo-600"/>
                        <span className="font-bold text-xs text-slate-700">إدارة قاعدة البيانات</span>
                      </div>
                      <div className="flex flex-col gap-1">
                        <button onClick={() => { setShowSystemMenu(false); handleBackupAll(); }} className="flex items-center gap-3 px-3 py-2.5 bg-indigo-50 hover:bg-indigo-100 transition-colors text-right rounded-xl text-indigo-700 font-bold text-xs">
                          <div className="bg-indigo-200/50 p-1.5 rounded-lg"><Download size={14} /></div> نسخ احتياطي
                        </button>
                        <button onClick={() => { setShowSystemMenu(false); restoreFileRef.current?.click(); }} className="flex items-center gap-3 px-3 py-2.5 bg-emerald-50 hover:bg-emerald-100 transition-colors text-right rounded-xl text-emerald-700 font-bold text-xs">
                          <div className="bg-emerald-200/50 p-1.5 rounded-lg"><Upload size={14} /></div> استعادة بيانات
                        </button>
                        <button onClick={() => { setShowSystemMenu(false); handleDeleteAllData(); }} className="flex items-center gap-3 px-3 py-2.5 bg-red-50 hover:bg-red-100 transition-colors text-right rounded-xl text-red-600 font-bold text-xs">
                          <div className="bg-red-200/50 p-1.5 rounded-lg"><Trash2 size={14} /></div> حذف جميع البيانات
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              ) : null
            } 
          />
        ) : activeTab === "budget" ? (
          <BudgetSection supabase={supabase} currentUser={currentUser} showToast={showToast} setConfirmDialog={setConfirmDialog} />
        ) : activeTab === "assets_new" ? (
          <AssetsSection supabase={supabase} currentUser={currentUser} showToast={showToast} setConfirmDialog={setConfirmDialog} />
        ) : activeTab === "admin_reports" ? (
          <AdminReportsSection supabase={supabase} currentUser={currentUser} showToast={showToast} setConfirmDialog={setConfirmDialog} />
        ) : (
          <DataTableTab key={activeTab} schemaId={activeTab} supabase={supabase} currentUser={currentUser} logAction={logAction} showToast={showToast} setConfirmDialog={setConfirmDialog} />
        )}
      </main>

      {toastMessage && (
        <div className={`fixed bottom-6 right-6 p-3 rounded-xl shadow-xl flex items-center gap-2 z-[9999] text-white font-bold animate-in slide-in-from-bottom-5 ${toastMessage.type === 'success' ? 'bg-emerald-500' : toastMessage.type === 'error' ? 'bg-red-500' : 'bg-blue-500'}`}>
          {toastMessage.type === 'success' ? <CheckCircle size={18} /> : toastMessage.type === 'error' ? <X size={18} /> : <div />}
          {toastMessage.msg}
        </div>
      )}

      {confirmDialog && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[10000] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl p-5 max-w-sm w-full shadow-2xl">
            <h3 className="text-lg font-black text-slate-800 mb-2">{confirmDialog.title}</h3>
            <p className="text-slate-600 mb-4 font-medium leading-relaxed text-sm">{confirmDialog.msg}</p>
            {confirmDialog.requirePassword && (
              <div className="mb-4">
                <label className="block text-xs font-bold text-slate-700 mb-1.5">كلمة المرور للتأكيد</label>
                <input type="password" id="confirm-dialog-pass" className="w-full border-2 border-slate-200 rounded-xl p-2.5 outline-none focus:border-red-500 text-sm" placeholder="أدخل كلمة المرور..." />
              </div>
            )}
            <div className="flex gap-2">
              <button 
                onClick={() => {
                  if (confirmDialog.requirePassword) {
                    const val = (document.getElementById("confirm-dialog-pass") as HTMLInputElement)?.value;
                    if (val !== "2026") {
                      showToast("كلمة المرور غير صحيحة", "error");
                      return;
                    }
                  }
                  confirmDialog.onConfirm();
                  setConfirmDialog(null);
                }} 
                className="flex-1 bg-red-500 hover:bg-red-600 text-white font-bold py-2.5 rounded-xl transition-colors text-sm"
              >
                تأكيد المتابعة
              </button>
              <button onClick={() => setConfirmDialog(null)} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 rounded-xl transition-colors text-sm">
                إلغاء
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ==================== ADMIN REPORTS INTEGRATION: COMPLETE ====================


// ==================== ADMIN REPORTS SECTION ====================

const ARABIC_MONTHS_LABELS = [
  "يناير","فبراير","مارس","أبريل","مايو","يونيو",
  "يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"
];

interface AdminReport {
  id: string;
  report_date: string | null;
  voucher_no: string;
  item_code: string;
  store_name: string;
  item_name: string;
  unit: string;
  quantity: number;
  cost_center: string;
  cost_center_description: string;
  segment: string;
  task: string;
  task_description: string;
  notes: string;
  price: number;
  total_value: number;
  report_month: string;
  report_year: string;
}

const AdminReportsSection = ({ supabase, currentUser, showToast, setConfirmDialog }: any) => {
  const isViewer = currentUser?.role === "viewer";
  const [records, setRecords] = useState<AdminReport[]>([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [editingRow, setEditingRow] = useState<AdminReport | null>(null);
  const [saving, setSaving] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [importProgress, setImportProgress] = useState<any>(null);
  const [searchText, setSearchText] = useState("");
  const [monthFilter, setMonthFilter] = useState("all");
  const [yearFilter, setYearFilter] = useState("all");
  const [storeFilter, setStoreFilter] = useState("all");
  const [taskFilter, setTaskFilter] = useState("all");
  const [form, setForm] = useState<any>({});
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isMobile = useIsMobile();
  const headerPadTop = isMobile ? 40 : 12;

  const TABLE = "admin_reports";

  const fetchAll = useCallback(async () => {
    setLoading(true);
    try {
      const data = await fetchAllPages(() => supabase.from(TABLE).select("*").order("report_date", { ascending: false }));
      setRecords(data);
    } catch (e: any) {
      showToast("خطأ في تحميل بيانات التقارير: " + e.message, "error");
    }
    setLoading(false);
  }, [supabase, showToast]);

  useEffect(() => { fetchAll(); }, [fetchAll]);

  // فلترة
  const filtered = useMemo(() => {
    let r = [...records];
    if (searchText) {
      const s = searchText.toLowerCase();
      r = r.filter(x =>
        x.item_name?.toLowerCase().includes(s) ||
        x.store_name?.toLowerCase().includes(s) ||
        x.voucher_no?.toLowerCase().includes(s) ||
        x.item_code?.toLowerCase().includes(s) ||
        x.cost_center_description?.toLowerCase().includes(s) ||
        x.task_description?.toLowerCase().includes(s)
      );
    }
    if (monthFilter !== "all") r = r.filter(x => x.report_month === monthFilter);
    if (yearFilter !== "all") r = r.filter(x => x.report_year === yearFilter);
    if (storeFilter !== "all") r = r.filter(x => x.store_name === storeFilter);
    if (taskFilter !== "all") r = r.filter(x => x.task === taskFilter);
    return r;
  }, [records, searchText, monthFilter, yearFilter, storeFilter, taskFilter]);

  // إحصائيات
  const stats = useMemo(() => {
    const totalValue = filtered.reduce((s, r) => s + (Number(r.total_value) || 0), 0);
    const totalQty = filtered.reduce((s, r) => s + (Number(r.quantity) || 0), 0);
    const uniqueItems = new Set(filtered.map(r => r.item_name)).size;
    const uniqueVouchers = new Set(filtered.map(r => r.voucher_no)).size;
    return { totalValue, totalQty, uniqueItems, uniqueVouchers, count: filtered.length };
  }, [filtered]);

  // بيانات الرسم البياني - المصروفات حسب المخزن
  const storeChartData = useMemo(() => {
    const m: Record<string, number> = {};
    filtered.forEach(r => {
      const k = r.store_name || "غير محدد";
      m[k] = (m[k] || 0) + (Number(r.total_value) || 0);
    });
    return Object.entries(m).map(([name, value]) => ({ name, value: Math.round(value) }))
      .sort((a, b) => b.value - a.value).slice(0, 8);
  }, [filtered]);

  // بيانات الرسم - المصروفات حسب Task
  const taskChartData = useMemo(() => {
    const m: Record<string, number> = {};
    filtered.forEach(r => {
      const k = r.task_description || r.task || "غير محدد";
      m[k] = (m[k] || 0) + (Number(r.total_value) || 0);
    });
    return Object.entries(m).map(([name, value]) => ({ name: name.length > 16 ? name.slice(0,14)+"…" : name, value: Math.round(value) }))
      .sort((a, b) => b.value - a.value).slice(0, 6);
  }, [filtered]);

  // قوائم الفلترة
  const storeOptions = useMemo(() => [...new Set(records.map(r => r.store_name).filter(Boolean))].sort(), [records]);
  const taskOptions = useMemo(() => [...new Set(records.map(r => r.task).filter(Boolean))].sort(), [records]);
  const yearOptions = useMemo(() => [...new Set(records.map(r => r.report_year).filter(Boolean))].sort().reverse(), [records]);

  const openAdd = () => {
    setForm({
      report_date: "", voucher_no: "", item_code: "", store_name: "",
      item_name: "", unit: "عدد", quantity: 0, cost_center: "",
      cost_center_description: "", segment: "", task: "", task_description: "",
      notes: "", price: 0, total_value: 0,
      report_month: ARABIC_MONTHS_LABELS[new Date().getMonth()],
      report_year: new Date().getFullYear().toString()
    });
    setEditingRow(null);
    setShowForm(true);
  };

  const openEdit = (row: AdminReport) => {
    setForm({ ...row });
    setEditingRow(row);
    setShowForm(true);
  };

  const save = async () => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    if (!form.item_name?.trim()) return showToast("اسم الصنف مطلوب", "error");
    setSaving(true);
    try {
      const data = {
        ...form,
        quantity: Number(form.quantity) || 0,
        price: Number(form.price) || 0,
        total_value: Number(form.total_value) || (Number(form.quantity) * Number(form.price)),
        updated_at: new Date().toISOString()
      };
      if (!editingRow) {
        delete data.id;
        const { error } = await supabase.from(TABLE).insert([{ ...data, created_by: currentUser?.id || null, created_at: new Date().toISOString() }]);
        if (error) throw error;
      } else {
        const { error } = await supabase.from(TABLE).update(data).eq("id", editingRow.id);
        if (error) throw error;
      }
      setShowForm(false);
      await fetchAll();
      showToast(editingRow ? "تم التعديل بنجاح" : "تمت الإضافة بنجاح", "success");
    } catch (e: any) {
      showToast("خطأ: " + e.message, "error");
    }
    setSaving(false);
  };

  const deleteRow = (id: string) => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    setConfirmDialog({
      title: "تأكيد الحذف",
      msg: "هل تريد حذف هذا السجل؟",
      onConfirm: async () => {
        const { error } = await supabase.from(TABLE).delete().eq("id", id);
        if (error) { showToast("خطأ في الحذف: " + error.message, "error"); return; }
        await fetchAll();
        showToast("تم الحذف بنجاح", "success");
      }
    });
  };

  const deleteAll = () => {
    if (isViewer) return showToast("ليس لديك صلاحية", "error");
    setConfirmDialog({
      title: "حذف جميع التقارير",
      msg: "هل تريد حذف جميع بيانات التقارير الإدارية؟ لا يمكن التراجع عن هذا الإجراء!",
      requirePassword: true,
      onConfirm: async () => {
        for (const r of records) {
          await supabase.from(TABLE).delete().eq("id", r.id);
        }
        await fetchAll();
        showToast("تم حذف جميع البيانات", "success");
      }
    });
  };

  // استيراد من Excel
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    setImportProgress({ total: 0, done: 0, errors: [], success: false });

    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true });

      // نقرأ كل الشيتات
      const toInsert: any[] = [];
      const errors: any[] = [];

      // نسأل المستخدم عن الشهر والسنة من اسم الملف
      const fileName = file.name;
      let detectedMonth = "";
      ARABIC_MONTHS_LABELS.forEach(m => {
        if (fileName.includes(m)) detectedMonth = m;
      });
      const monthMap: Record<string, string> = {
        "يناير": "يناير", "فبراير": "فبراير", "مارس": "مارس", "أبريل": "أبريل",
        "ابريل": "أبريل", "مايو": "مايو", "يونيو": "يونيو", "يوليو": "يوليو",
        "أغسطس": "أغسطس", "اغسطس": "أغسطس", "سبتمبر": "سبتمبر", "أكتوبر": "أكتوبر",
        "اكتوبر": "أكتوبر", "نوفمبر": "نوفمبر", "ديسمبر": "ديسمبر"
      };
      // تحقق من اسم الملف
      Object.keys(monthMap).forEach(k => {
        if (fileName.includes(k)) detectedMonth = monthMap[k];
      });
      if (!detectedMonth) detectedMonth = ARABIC_MONTHS_LABELS[new Date().getMonth()];

      const yearMatch = fileName.match(/20\d{2}/);
      const detectedYear = yearMatch ? yearMatch[0] : new Date().getFullYear().toString();

      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as any[][];

        // البحث عن صف الهيدر
        let headerIdx = -1;
        const headerKeywords = ["اسم الصنف", "Item", "كمية الصرف", "السعر", "القيمة"];
        for (let i = 0; i < Math.min(raw.length, 10); i++) {
          const rowVals = raw[i].map((v: any) => String(v ?? "").trim());
          const matches = headerKeywords.filter(k => rowVals.some(rv => rv.includes(k))).length;
          if (matches >= 2) { headerIdx = i; break; }
        }
        if (headerIdx < 0) continue;

        const headers = (raw[headerIdx] || []).map((h: any) => String(h ?? "").trim());

        const colIdx = (names: string[]) => {
          for (const n of names) {
            const i = headers.findIndex(h => h.includes(n) || n.includes(h));
            if (i >= 0) return i;
          }
          return -1;
        };

        const dateIdx = colIdx(["التاريخ", "Date"]);
        const voucherIdx = colIdx(["رقم اذن", "رقم الاذن", "Voucher"]);
        const itemCodeIdx = colIdx(["كود الصنف", "Item code", "الكود"]);
        const storeIdx = colIdx(["اسم المخزن", "المخزن", "Store"]);
        const itemNameIdx = colIdx(["اسم الصنف", "الصنف", "Item"]);
        const unitIdx = colIdx(["الوحدة", "Unit"]);
        const qtyIdx = colIdx(["كمية الصرف", "الكمية", "Qty", "Quantity"]);
        const costCenterIdx = colIdx(["Cost center", "cost center"]);
        const costCenterDescIdx = colIdx(["Cost center Description", "Description"]);
        const segmentIdx = colIdx(["Segment"]);
        const taskIdx = colIdx(["Task"]);
        const taskDescIdx = colIdx(["Task Description"]);
        const notesIdx = colIdx(["ملاحظات", "Notes"]);
        const priceIdx = colIdx(["السعر", "Price"]);
        const valueIdx = colIdx(["القيمة", "Value", "Total"]);

        for (let i = headerIdx + 1; i < raw.length; i++) {
          const row = raw[i];
          if (!row) continue;

          const itemVal = String(row[itemNameIdx] ?? "").trim();
          if (!itemVal || itemVal === "اسم الصنف" || itemVal === "Item") continue;

          const getStr = (idx: number) => idx >= 0 ? String(row[idx] ?? "").trim() : "";
          const getNum = (idx: number) => {
            if (idx < 0 || row[idx] == null) return 0;
            const n = parseFloat(arabicToEnglish(String(row[idx])));
            return isNaN(n) ? 0 : n;
          };

          // تحويل التاريخ
          let reportDate: string | null = null;
          if (dateIdx >= 0 && row[dateIdx] != null) {
            reportDate = parseImportDate(row[dateIdx]);
          }

          const qty = getNum(qtyIdx);
          const price = getNum(priceIdx);
          const totalValue = getNum(valueIdx) || (qty * price);

          toInsert.push({
            report_date: reportDate,
            voucher_no: getStr(voucherIdx),
            item_code: getStr(itemCodeIdx),
            store_name: getStr(storeIdx),
            item_name: itemVal,
            unit: getStr(unitIdx) || "عدد",
            quantity: qty,
            cost_center: getStr(costCenterIdx),
            cost_center_description: getStr(costCenterDescIdx),
            segment: getStr(segmentIdx),
            task: getStr(taskIdx),
            task_description: getStr(taskDescIdx),
            notes: getStr(notesIdx),
            price,
            total_value: totalValue,
            report_month: detectedMonth,
            report_year: detectedYear,
            created_by: currentUser?.id || null,
            created_at: new Date().toISOString(),
            updated_at: new Date().toISOString()
          });
        }
      }

      if (toInsert.length === 0) {
        setImportProgress({ total: 0, done: 0, errors: [{ msg: "لم يتم العثور على بيانات صالحة. تأكد من أن الملف يحتوي على الأعمدة المطلوبة." }], success: false });
        setUploading(false);
        if (fileInputRef.current) fileInputRef.current.value = "";
        return;
      }

      // ── Upsert: skip duplicates, update changed, insert new ──
      // Unique key: voucher_no + item_name (fallback: store_name + item_name + report_date)
      let inserted = 0, updated = 0, done = 0;
      for (const rec of toInsert) {
        const existing = records.find((r: any) => {
          if (rec.voucher_no && r.voucher_no)
            return r.voucher_no === rec.voucher_no && r.item_name === rec.item_name;
          return r.store_name === rec.store_name &&
                 r.item_name === rec.item_name &&
                 r.report_date === rec.report_date;
        });
        if (existing) {
          // Check if anything actually changed
          const changed = existing.quantity !== rec.quantity ||
                          existing.total_value !== rec.total_value ||
                          existing.price !== rec.price;
          if (changed) {
            const { error } = await supabase.from(TABLE).update({ ...rec, updated_at: new Date().toISOString() }).eq("id", existing.id);
            if (!error) updated++;
            else errors.push({ msg: `تحديث "${rec.item_name}": ${error.message}` });
          }
          // else identical → skip (no duplicate)
        } else {
          const { error } = await supabase.from(TABLE).insert([rec]);
          if (!error) inserted++;
          else errors.push({ msg: `إضافة "${rec.item_name}": ${error.message}` });
        }
        done++;
        if (done % 10 === 0 || done === toInsert.length)
          setImportProgress({ total: toInsert.length, done, errors, success: false });
      }

      setImportProgress({ total: toInsert.length, done, errors, success: errors.length === 0, inserted, updated, month: detectedMonth, year: detectedYear });
      await fetchAll();
    } catch (e: any) {
      console.error("Import error:", e);
      setImportProgress({ total: 0, done: 0, errors: [{ msg: e.message }], success: false });
    }
    setUploading(false);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  // تصدير Excel
  const exportToExcel = () => {
    if (!filtered.length) return showToast("لا توجد بيانات للتصدير", "error");
    const exportData = filtered.map(r => ({
      "التاريخ": r.report_date ? String(r.report_date).split("T")[0] : "",
      "رقم اذن الصرف": r.voucher_no,
      "كود الصنف": r.item_code,
      "اسم المخزن": r.store_name,
      "اسم الصنف": r.item_name,
      "الوحدة": r.unit,
      "كمية الصرف": r.quantity,
      "Cost center": r.cost_center,
      "Cost center Description": r.cost_center_description,
      "Segment": r.segment,
      "Task": r.task,
      "Task Description": r.task_description,
      "ملاحظات": r.notes,
      "السعر": r.price,
      "القيمة": r.total_value,
      "الشهر": r.report_month,
      "السنة": r.report_year
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "التقارير الإدارية");
    XLSX.writeFile(wb, `التقارير_الإدارية_${monthFilter !== "all" ? monthFilter : "الكل"}_${new Date().toISOString().split("T")[0]}.xlsx`);
    showToast("تم التصدير بنجاح", "success");
  };

  const downloadTemplate = () => {
    const templateData = [{
      "التاريخ": "2025-04-01",
      "رقم اذن الصرف": "36698",
      "كود الصنف": "2011695",
      "اسم المخزن": "ميكانيكا",
      "اسم الصنف": "مثال صنف",
      "الوحدة": "عدد",
      "كمية الصرف": 1,
      "Cost center": "gen",
      "Cost center Description": "General",
      "Segment": "General",
      "Task": "210",
      "Task Description": "اصلاحات ادارى",
      "ملاحظات": "",
      "السعر": 20.52,
      "القيمة": 20.52
    }];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "المصروفات");
    XLSX.writeFile(wb, "نموذج_التقارير_الإدارية.xlsx");
  };

  const PIE_COLORS = ["#6366f1","#f59e0b","#10b981","#f43f5e","#8b5cf6","#14b8a6","#3b82f6","#ec4899"];

  return (
    <div className="space-y-3" dir="rtl">
      {/* ======== الهيدر ======== */}
      <div style={{ position:"sticky", top: -headerPadTop, zIndex:50, background:"rgba(248,250,252,0.98)", backdropFilter:"blur(12px)", borderBottom:"1px solid rgba(226,232,240,0.9)", boxShadow:"0 10px 30px -12px rgba(0,0,0,0.08)", paddingTop: headerPadTop, paddingBottom:"4px", marginLeft:"-4px", marginRight:"-4px", paddingLeft:"4px", paddingRight:"4px" }}>
        <div className="bg-gradient-to-l from-cyan-700 to-teal-800 text-white px-4 py-3 rounded-xl flex items-center justify-between flex-wrap gap-2" style={{ boxShadow:"0 10px 30px -12px rgba(13,148,136,0.5)" }}>
          <div>
            <h2 className="text-xl font-black flex items-center gap-2">📋 التقارير الإدارية</h2>
            <p className="text-cyan-100 text-xs font-bold mt-0.5">تقارير المصروفات الشهرية لمزارع لينة</p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <button onClick={downloadTemplate}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.18)", color:"white", border:"1px solid rgba(255,255,255,0.3)", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <FileDown size={13} /> نموذج Excel
            </button>
            {!isViewer && (
              <button onClick={() => fileInputRef.current?.click()} disabled={uploading}
                style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.18)", color:"white", border:"1px solid rgba(255,255,255,0.3)", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
                <Upload size={13} /> {uploading ? "جاري..." : "استيراد Excel"}
              </button>
            )}
            <button onClick={exportToExcel}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"rgba(255,255,255,0.9)", color:"#134e4a", border:"none", borderRadius:"8px", padding:"6px 10px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Download size={13} /> تصدير
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{ display:"none" }} />
          </div>
        </div>

        {/* بطاقات الإحصاء */}
        <div style={{ marginTop:"6px", display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:"6px" }}>
          <div style={{ background:"linear-gradient(135deg,#f0fdfa,#ccfbf1)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #5eead4", textAlign:"center" }}>
            <p style={{ margin:0, color:"#0f766e", fontWeight:"800", fontSize:"9px" }}>إجمالي السجلات</p>
            <h3 style={{ margin:"2px 0 0", color:"#0d9488", fontSize:"14px", fontWeight:"900" }}>{stats.count.toLocaleString("en-US")}</h3>
          </div>
          <div style={{ background:"linear-gradient(135deg,#fffbeb,#fef3c7)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #fcd34d", textAlign:"center" }}>
            <p style={{ margin:0, color:"#92400e", fontWeight:"800", fontSize:"9px" }}>إجمالي القيمة</p>
            <h3 style={{ margin:"2px 0 0", color:"#d97706", fontSize:"13px", fontWeight:"900" }}>{stats.totalValue.toLocaleString("en-US", { maximumFractionDigits: 0 })} ج</h3>
          </div>
          <div style={{ background:"linear-gradient(135deg,#f5f3ff,#ede9fe)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #c4b5fd", textAlign:"center" }}>
            <p style={{ margin:0, color:"#4c1d95", fontWeight:"800", fontSize:"9px" }}>عدد الأصناف</p>
            <h3 style={{ margin:"2px 0 0", color:"#6d28d9", fontSize:"14px", fontWeight:"900" }}>{stats.uniqueItems.toLocaleString("en-US")}</h3>
          </div>
          <div style={{ background:"linear-gradient(135deg,#eff6ff,#dbeafe)", padding:"6px 8px", borderRadius:"8px", border:"1px solid #93c5fd", textAlign:"center" }}>
            <p style={{ margin:0, color:"#1e3a8a", fontWeight:"800", fontSize:"9px" }}>عدد أذونات الصرف</p>
            <h3 style={{ margin:"2px 0 0", color:"#1d4ed8", fontSize:"14px", fontWeight:"900" }}>{stats.uniqueVouchers.toLocaleString("en-US")}</h3>
          </div>
        </div>

        {/* شريط الشهور */}
        <div style={{ marginTop:"8px", display:"flex", gap:"4px", flexWrap:"wrap", background:"rgba(255,255,255,0.96)", backdropFilter:"blur(8px)", padding:"8px", borderRadius:"10px", border:"1px solid #e2e8f0", boxShadow:"0 4px 12px rgba(0,0,0,0.04)" }}>
          <button onClick={() => setMonthFilter("all")}
            style={{ padding:"5px 10px", borderRadius:"6px", fontWeight:"800", fontSize:"11px", cursor:"pointer", border:"none",
              background: monthFilter === "all" ? "linear-gradient(90deg,#0f766e,#0d9488)" : "#f1f5f9",
              color: monthFilter === "all" ? "white" : "#64748b",
              boxShadow: monthFilter === "all" ? "0 3px 10px rgba(13,148,136,0.35)" : "none"
            }}>
            الكل {monthFilter === "all" ? `(${records.length})` : ""}
          </button>
          {ARABIC_MONTHS_LABELS.map(m => {
            const cnt = records.filter(r => r.report_month === m).length;
            if (cnt === 0) return null;
            return (
              <button key={m} onClick={() => setMonthFilter(m)}
                style={{ padding:"5px 10px", borderRadius:"6px", fontWeight:"800", fontSize:"11px", cursor:"pointer", border:"none",
                  background: monthFilter === m ? "linear-gradient(90deg,#0f766e,#0d9488)" : "#f1f5f9",
                  color: monthFilter === m ? "white" : "#64748b",
                  boxShadow: monthFilter === m ? "0 3px 10px rgba(13,148,136,0.35)" : "none"
                }}>
                {m} <span style={{ opacity:0.75, fontSize:"10px" }}>({cnt})</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* رسائل الاستيراد */}
      <div style={{ background:"white", padding:"12px", borderRadius:"12px", border:"1px solid #e2e8f0" }}>
        <h3 style={{ margin:"0 0 8px", fontWeight:"800", fontSize:"12px", color:"#4c1d95" }}>
          📊 الكميات الشهرية{activeDept !== "all" ? ` — ${activeDept}` : " — جميع الجهات"}
        </h3>
        <div style={{ height:"140px" }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={monthlyTotalsForDept} margin={{ top:0, right:0, left:-20, bottom:0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
              <XAxis dataKey="name" tick={{ fontSize:9, fill:"#94a3b8" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize:9, fill:"#94a3b8" }} axisLine={false} tickLine={false} />
              <RechartsTooltip
                formatter={(v: any) => [Number(v).toLocaleString("en-US"), "الكمية"]}
                contentStyle={{ borderRadius:"8px", border:"none", boxShadow:"0 4px 20px rgba(0,0,0,0.1)", fontSize:"11px" }}
              />
              <Bar dataKey="total" radius={[6,6,0,0]} barSize={16} fill="#7c3aed" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {importProgress && (
        <div style={{ background:"white", borderRadius:"10px", padding:"12px", border:`2px solid ${importProgress.success ? "#22c55e" : "#3b82f6"}` }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:"6px" }}>
            <p style={{ margin:0, fontWeight:"900", fontSize:"13px" }}>
              {importProgress.success
                ? `✅ اكتمل الاستيراد — تم رفع ${importProgress.done} سجل لشهر ${importProgress.month || ""} ${importProgress.year || ""}`
                : "📊 جاري الاستيراد..."}
            </p>
            <button onClick={() => setImportProgress(null)} style={{ background:"none", border:"none", cursor:"pointer", fontSize:"16px" }}>✕</button>
          </div>
          <div style={{ background:"#e2e8f0", borderRadius:"99px", height:"6px", overflow:"hidden" }}>
            <div style={{ height:"100%", width:`${importProgress.total ? (importProgress.done/importProgress.total)*100 : 0}%`, background: importProgress.success ? "#22c55e" : "#4f46e5", borderRadius:"99px", transition:"width 0.3s" }} />
          </div>
          <p style={{ margin:"4px 0 0", fontSize:"11px", color:"#64748b", fontWeight:"700" }}>تم: {importProgress.done} من {importProgress.total} · {(importProgress as any).inserted > 0 ? `✅ جديد: ${(importProgress as any).inserted}` : ""} · {(importProgress as any).updated > 0 ? `♻️ محدّث: ${(importProgress as any).updated}` : ""}</p>
          {importProgress.errors?.length > 0 && (
            <div style={{ background:"#fef9ec", borderRadius:"6px", padding:"8px", marginTop:"6px", fontSize:"11px", color:"#78350f", maxHeight:"100px", overflowY:"auto" }}>
              {importProgress.errors.map((e: any, i: number) => <div key={i}>• {e.msg}</div>)}
            </div>
          )}
        </div>
      )}

      {/* رسوم بيانية */}
      {filtered.length > 0 && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
          <div style={{ background:"white", borderRadius:"12px", padding:"12px", border:"1px solid #e2e8f0" }}>
            <h3 style={{ margin:"0 0 8px", fontWeight:"800", fontSize:"12px", color:"#0f766e" }}>📊 المصروفات حسب المخزن</h3>
            <div style={{ height:"180px" }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={storeChartData} margin={{ top:0, right:0, left:-15, bottom:0 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis dataKey="name" tick={{ fontSize:9, fill:"#64748b" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize:9, fill:"#64748b" }} axisLine={false} tickLine={false} />
                  <RechartsTooltip formatter={(v: any) => [Number(v).toLocaleString("en-US") + " ج", "القيمة"]} contentStyle={{ borderRadius:"8px", border:"none", boxShadow:"0 4px 20px rgba(0,0,0,0.1)", fontSize:"11px" }} />
                  <Bar dataKey="value" radius={[6,6,0,0]} barSize={20}>
                    {storeChartData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div style={{ background:"white", borderRadius:"12px", padding:"12px", border:"1px solid #e2e8f0" }}>
            <h3 style={{ margin:"0 0 8px", fontWeight:"800", fontSize:"12px", color:"#0f766e" }}>🎯 المصروفات حسب Task</h3>
            <div style={{ height:"180px" }}>
              <ResponsiveContainer width="100%" height="100%">
                <RePieChart>
                  <Pie data={taskChartData} cx="50%" cy="50%" innerRadius={45} outerRadius={70} dataKey="value" paddingAngle={3}>
                    {taskChartData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                  </Pie>
                  <RechartsTooltip formatter={(v: any) => [Number(v).toLocaleString("en-US") + " ج", "القيمة"]} contentStyle={{ borderRadius:"8px", border:"none", boxShadow:"0 4px 20px rgba(0,0,0,0.1)", fontSize:"11px" }} />
                  <Legend wrapperStyle={{ fontSize:"10px", fontWeight:700 }} />
                </RePieChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      )}

      {/* شريط الفلترة */}
      <div style={{ background:"white", borderRadius:"10px", padding:"10px 12px", border:"1px solid #e2e8f0", display:"flex", gap:"8px", flexWrap:"wrap", alignItems:"center" }}>
        <div style={{ flex:1, minWidth:"180px", position:"relative" }}>
          <Search size={14} style={{ position:"absolute", right:"8px", top:"50%", transform:"translateY(-50%)", color:"#94a3b8" }} />
          <input type="text" placeholder="بحث في الصنف، المخزن، الإذن..." value={searchText} onChange={e => setSearchText(e.target.value)}
            style={{ width:"100%", padding:"8px 28px 8px 10px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontSize:"12px", fontFamily:"Cairo,sans-serif" }} />
        </div>

        <select value={yearFilter} onChange={e => setYearFilter(e.target.value)}
          style={{ padding:"8px 10px", border:"1px solid #e2e8f0", borderRadius:"6px", fontSize:"12px", outline:"none", fontFamily:"Cairo,sans-serif" }}>
          <option value="all">جميع السنوات</option>
          {yearOptions.map(y => <option key={y} value={y}>{y}</option>)}
        </select>
        <select value={storeFilter} onChange={e => setStoreFilter(e.target.value)}
          style={{ padding:"8px 10px", border:"1px solid #e2e8f0", borderRadius:"6px", fontSize:"12px", outline:"none", fontFamily:"Cairo,sans-serif" }}>
          <option value="all">جميع المخازن</option>
          {storeOptions.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <select value={taskFilter} onChange={e => setTaskFilter(e.target.value)}
          style={{ padding:"8px 10px", border:"1px solid #e2e8f0", borderRadius:"6px", fontSize:"12px", outline:"none", fontFamily:"Cairo,sans-serif" }}>
          <option value="all">جميع الـ Tasks</option>
          {taskOptions.map(t => <option key={t} value={t}>{t}</option>)}
        </select>
        {!isViewer && (
          <>
            <button onClick={openAdd}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"#0d9488", color:"white", border:"none", borderRadius:"8px", padding:"8px 12px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Plus size={14} /> إضافة سجل
            </button>
            <button onClick={deleteAll}
              style={{ display:"flex", alignItems:"center", gap:"5px", background:"#fee2e2", color:"#dc2626", border:"none", borderRadius:"8px", padding:"8px 12px", fontWeight:"800", cursor:"pointer", fontSize:"12px" }}>
              <Trash2 size={14} /> حذف الكل
            </button>
          </>
        )}
      </div>

      {/* الجدول */}
      {loading ? (
        <div style={{ textAlign:"center", padding:"40px" }}><Loader2 className="animate-spin text-teal-500" size={32} /></div>
      ) : (
        <div style={{ background:"white", borderRadius:"12px", border:"1px solid #e2e8f0", overflow:"auto", maxHeight:"calc(100vh - 300px)", boxShadow:"0 2px 10px rgba(0,0,0,0.03)" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:"12px", minWidth:"1200px" }}>
            <thead style={{ background:"linear-gradient(90deg,#0f766e,#0d9488)", color:"white", position:"sticky", top:0, zIndex:5 }}>
              <tr>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px", whiteSpace:"nowrap" }}>التاريخ</th>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px", whiteSpace:"nowrap" }}>رقم الإذن</th>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px", whiteSpace:"nowrap" }}>الكود</th>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>المخزن</th>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>اسم الصنف</th>
                <th style={{ padding:"6px 8px", textAlign:"center", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>الوحدة</th>
                <th style={{ padding:"6px 8px", textAlign:"center", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>الكمية</th>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>Cost Center</th>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>Task</th>
                <th style={{ padding:"6px 8px", textAlign:"right", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>Task Description</th>
                <th style={{ padding:"6px 8px", textAlign:"center", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>السعر</th>
                <th style={{ padding:"6px 8px", textAlign:"center", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px", background:"rgba(0,0,0,0.15)" }}>القيمة</th>
                <th style={{ padding:"6px 8px", textAlign:"center", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>الشهر</th>
                {!isViewer && <th style={{ padding:"6px 8px", textAlign:"center", border:"1px solid #0f766e", fontWeight:"800", fontSize:"11px" }}>إجراءات</th>}
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={14} style={{ textAlign:"center", padding:"40px", color:"#94a3b8", fontWeight:"700", fontSize:"14px" }}>لا توجد بيانات — قم برفع ملف التقرير الشهري أو الإضافة اليدوية</td></tr>
              ) : filtered.map((row, i) => {
                const bg = i % 2 === 0 ? "white" : "#f0fdfa";
                return (
                  <tr key={row.id} style={{ background: bg, transition:"background 0.15s" }}
                    onMouseEnter={e => (e.currentTarget.style.background = "#ccfbf1")}
                    onMouseLeave={e => (e.currentTarget.style.background = bg)}>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", color:"#334155", fontSize:"11px", whiteSpace:"nowrap" }}>
                      {row.report_date ? formatDate(row.report_date) : "-"}
                    </td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", color:"#6366f1", fontWeight:"700", fontSize:"11px", fontFamily:"monospace" }}>{row.voucher_no || "-"}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", color:"#64748b", fontSize:"11px", fontFamily:"monospace" }}>{row.item_code || "-"}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", fontSize:"11px" }}>
                      <span style={{ background:"#f0f9ff", color:"#0369a1", padding:"2px 6px", borderRadius:"12px", fontSize:"10px", fontWeight:"700" }}>{row.store_name || "-"}</span>
                    </td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", fontWeight:"700", color:"#1e293b", fontSize:"12px", maxWidth:"200px" }}>{row.item_name}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", textAlign:"center", color:"#64748b", fontSize:"11px" }}>{row.unit || "-"}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", textAlign:"center", fontWeight:"700", color:"#0d9488", fontSize:"12px" }}>{Number(row.quantity) || "-"}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", color:"#64748b", fontSize:"10px" }}>{row.cost_center || "-"}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", color:"#64748b", fontSize:"11px" }}>
                      <span style={{ background:"#fef3c7", color:"#92400e", padding:"2px 5px", borderRadius:"8px", fontSize:"10px", fontWeight:"700" }}>{row.task || "-"}</span>
                    </td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", color:"#475569", fontSize:"11px", maxWidth:"160px" }}>{row.task_description || "-"}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", textAlign:"center", color:"#15803d", fontWeight:"700", fontSize:"11px" }}>{Number(row.price)?.toLocaleString("en-US", { maximumFractionDigits:2 })}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", textAlign:"center", fontWeight:"900", color:"#1d4ed8", fontSize:"12px", background:"#eff6ff" }}>{Number(row.total_value)?.toLocaleString("en-US", { maximumFractionDigits:2 })}</td>
                    <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", textAlign:"center" }}>
                      <span style={{ background:"#f0fdfa", color:"#0f766e", padding:"2px 8px", borderRadius:"12px", fontSize:"10px", fontWeight:"700", whiteSpace:"nowrap" }}>{row.report_month} {row.report_year}</span>
                    </td>
                    {!isViewer && (
                      <td style={{ padding:"5px 8px", border:"1px solid #e2e8f0", textAlign:"center" }}>
                        <div style={{ display:"flex", gap:"4px", justifyContent:"center" }}>
                          <button onClick={() => openEdit(row)} style={{ padding:"4px 8px", background:"#eff6ff", color:"#0284c7", border:"none", borderRadius:"6px", cursor:"pointer", fontSize:"12px" }}>✏️</button>
                          <button onClick={() => deleteRow(row.id)} style={{ padding:"4px 8px", background:"#fff1f2", color:"#dc2626", border:"none", borderRadius:"6px", cursor:"pointer", fontSize:"12px" }}>🗑️</button>
                        </div>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* نموذج الإضافة/التعديل */}
      {showForm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", backdropFilter:"blur(6px)", display:"flex", alignItems:"center", justifyContent:"center", padding:"16px", zIndex:9999 }}
          onClick={() => setShowForm(false)}>
          <div style={{ background:"white", borderRadius:"16px", width:"100%", maxWidth:"760px", padding:"20px", boxShadow:"0 32px 80px rgba(0,0,0,0.25)", maxHeight:"90vh", overflowY:"auto" }}
            dir="rtl" onClick={e => e.stopPropagation()}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"16px" }}>
              <h3 style={{ margin:0, fontWeight:"900", fontSize:"16px", color:"#0f766e" }}>{editingRow ? "✏️ تعديل سجل" : "➕ إضافة سجل جديد"}</h3>
              <button onClick={() => setShowForm(false)} style={{ border:"1px solid #e2e8f0", borderRadius:"6px", padding:"5px 10px", cursor:"pointer", background:"white" }}>✕</button>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:"10px", marginBottom:"12px" }}>
              {[
                { key:"report_date", label:"التاريخ", type:"date" },
                { key:"voucher_no", label:"رقم اذن الصرف", type:"text" },
                { key:"item_code", label:"كود الصنف", type:"text" },
                { key:"store_name", label:"اسم المخزن", type:"text" },
                { key:"item_name", label:"اسم الصنف *", type:"text", span:2 },
                { key:"unit", label:"الوحدة", type:"text" },
                { key:"quantity", label:"كمية الصرف", type:"number" },
                { key:"price", label:"السعر", type:"number" },
                { key:"total_value", label:"القيمة", type:"number" },
                { key:"cost_center", label:"Cost Center", type:"text" },
                { key:"cost_center_description", label:"Cost Center Description", type:"text", span:2 },
                { key:"segment", label:"Segment", type:"text" },
                { key:"task", label:"Task", type:"text" },
                { key:"task_description", label:"Task Description", type:"text" },
                { key:"notes", label:"ملاحظات", type:"text", span:3 },
              ].map((f: any) => (
                <div key={f.key} style={{ gridColumn: f.span ? `span ${f.span}` : "auto" }}>
                  <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>{f.label}</label>
                  <input type={f.type} value={form[f.key] ?? ""} onChange={e => setForm({ ...form, [f.key]: f.type === "number" ? e.target.value : e.target.value })}
                    style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box", fontSize:"12px" }} />
                </div>
              ))}
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>الشهر</label>
                <select value={form.report_month || ""} onChange={e => setForm({ ...form, report_month: e.target.value })}
                  style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box", fontSize:"12px" }}>
                  {ARABIC_MONTHS_LABELS.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>
              <div>
                <label style={{ fontSize:"11px", fontWeight:"700", color:"#64748b", display:"block", marginBottom:"3px" }}>السنة</label>
                <select value={form.report_year || ""} onChange={e => setForm({ ...form, report_year: e.target.value })}
                  style={{ width:"100%", padding:"8px", border:"1px solid #e2e8f0", borderRadius:"6px", outline:"none", fontFamily:"Cairo,sans-serif", boxSizing:"border-box", fontSize:"12px" }}>
                  {years.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:"10px", marginTop:"12px" }}>
              <button onClick={() => setShowForm(false)} style={{ padding:"10px", background:"#f1f5f9", border:"none", borderRadius:"6px", fontWeight:"900", cursor:"pointer", fontFamily:"Cairo,sans-serif", fontSize:"13px" }}>إلغاء</button>
              <button onClick={save} disabled={saving} style={{ padding:"10px", background:"#0d9488", color:"white", border:"none", borderRadius:"6px", fontWeight:"900", cursor:"pointer", fontFamily:"Cairo,sans-serif", fontSize:"13px" }}>
                {saving ? "جاري..." : "💾 حفظ"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
