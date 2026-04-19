import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Chart as ChartJS,
  ArcElement,
  Tooltip,
  Legend,
  CategoryScale,
  LinearScale,
  BarElement,
} from "chart.js";
import { Doughnut, Bar } from "react-chartjs-2";

ChartJS.register(ArcElement, Tooltip, Legend, CategoryScale, LinearScale, BarElement);

// ── Varsayılan kategoriler ──────────────────────────────────────────────────
const DEFAULT_CATS = [
  { id: 1, name: "Market / Süpermarket", color: "#1a7a4a", keywords: ["migros","bim","a101","carrefour","şok","market","süpermarket","metro"] },
  { id: 2, name: "Yemek & Restoran",     color: "#e07b00", keywords: ["restoran","burger","döner","pizza","cafe","kahve","yemek","köfte","mangal"] },
  { id: 3, name: "Ulaşım",               color: "#3266ad", keywords: ["uber","taksi","metrobüs","istanbul kart","istanbulkart","dolmuş"] },
  { id: 4, name: "Fatura & Abonelik",    color: "#9b3ab5", keywords: ["netflix","spotify","türk telekom","vodafone","turkcell","internet","elektrik","doğalgaz","fatura"] },
  { id: 5, name: "Giyim & Alışveriş",   color: "#c0392b", keywords: ["zara","h&m","koton","lcw","mango","giyim","kıyafet","ayakkabı","trendyol"] },
  { id: 6, name: "Sağlık",               color: "#0891b2", keywords: ["eczane","hastane","doktor","klinik","muayene","ilaç","sağlık"] },
  { id: 7, name: "Eğlence & Kültür",    color: "#d97706", keywords: ["sinema","konser","tiyatro","müze","spor","fitness","gym"] },
  { id: 8, name: "Diğer",               color: "#888780", keywords: [] },
];

const DEMO = [
  { date: "2025-03-01", desc: "Migros AVM",           amount: 847  },
  { date: "2025-03-02", desc: "İstanbul Kart yükleme",amount: 200  },
  { date: "2025-03-03", desc: "Netflix abonelik",      amount: 149  },
  { date: "2025-03-04", desc: "Zara giyim",            amount: 1250 },
  { date: "2025-03-05", desc: "Uber taksi",            amount: 85   },
  { date: "2025-03-06", desc: "BIM market",            amount: 412  },
  { date: "2025-03-07", desc: "Spotify müzik",         amount: 59   },
  { date: "2025-03-08", desc: "Burger King yemek",     amount: 310  },
  { date: "2025-03-10", desc: "Eczane ilaç",           amount: 275  },
  { date: "2025-03-11", desc: "Carrefour süpermarket", amount: 1120 },
  { date: "2025-03-12", desc: "Sinema bileti",         amount: 220  },
  { date: "2025-03-13", desc: "Elektrik faturası",     amount: 480  },
  { date: "2025-03-14", desc: "Köfteci Ramiz",         amount: 185  },
  { date: "2025-03-15", desc: "H&M kıyafet",           amount: 890  },
  { date: "2025-03-17", desc: "A101 market",           amount: 560  },
  { date: "2025-03-18", desc: "Taksi ücreti",          amount: 145  },
  { date: "2025-03-19", desc: "Doktor muayene",        amount: 400  },
  { date: "2025-03-21", desc: "Mangal restoranı",      amount: 650  },
  { date: "2025-03-23", desc: "İnternet faturası",     amount: 299  },
  { date: "2025-03-24", desc: "Koton alışveriş",       amount: 475  },
  { date: "2025-03-25", desc: "BIM market",            amount: 320  },
  { date: "2025-03-26", desc: "Cafe kahve",            amount: 120  },
  { date: "2025-02-03", desc: "Migros market",         amount: 730  },
  { date: "2025-02-05", desc: "Restoran yemek",        amount: 420  },
  { date: "2025-02-10", desc: "Zara giyim alışveriş",  amount: 980  },
  { date: "2025-02-15", desc: "Turkcell fatura",       amount: 420  },
  { date: "2025-02-20", desc: "Uber sürüş",            amount: 210  },
  { date: "2025-02-22", desc: "Gym fitness üyelik",    amount: 350  },
  { date: "2025-01-05", desc: "Migros haftalık",       amount: 900  },
  { date: "2025-01-12", desc: "Döner yemek",           amount: 180  },
  { date: "2025-01-18", desc: "Elektrik doğalgaz fatura", amount: 750 },
  { date: "2025-01-25", desc: "H&M alışveriş",         amount: 1100 },
];

// ── Yardımcı fonksiyonlar ───────────────────────────────────────────────────
function categorize(desc, cats) {
  const d = desc.toLowerCase();
  for (const cat of cats) {
    if (cat.name === "Diğer") continue;
    if (cat.keywords.some((k) => k && d.includes(k.toLowerCase()))) return cat.name;
  }
  return "Diğer";
}

function getColor(name, cats) {
  return cats.find((c) => c.name === name)?.color || "#888";
}

// ── Ana bileşen ─────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("upload");
  const [cats, setCats] = useState(JSON.parse(JSON.stringify(DEFAULT_CATS)));
  const [rawTx, setRawTx] = useState([]);
  const [tx, setTx] = useState([]);
  const [apiKey] = useState(import.meta.env.VITE_GEMINI_KEY || "");
  const [aiText, setAiText] = useState("");
  const [aiLoading, setAiLoading] = useState(false);
  const [filter, setFilter] = useState("");
  const [drag, setDrag] = useState(false);
  const fileRef = useRef();

  // Kategori değişince yeniden sınıflandır
  useEffect(() => {
    if (rawTx.length) setTx(rawTx.map((t) => ({ ...t, cat: categorize(t.desc, cats) })));
  }, [cats, rawTx]);

  function loadData(rows) {
    setRawTx(rows);
    setTx(rows.map((t) => ({ ...t, cat: categorize(t.desc, cats) })));
    setTab("analysis");
  }

  async function handleFile(file) {
    if (!file) return;
    if (file.name.match(/\.xlsx?$/i)) {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length < 3) continue;
        const desc = String(r[1] || "").trim();
        const amount = Math.abs(parseFloat(String(r[2] || "").replace(",", ".")));
        if (desc && !isNaN(amount) && amount > 0)
          parsed.push({ date: String(r[0] || ""), desc, amount });
      }
      if (parsed.length) loadData(parsed);
      else alert("Dosya okunamadı. Sütun sırası: Tarih, Açıklama, Tutar olmalı.");
    } else {
      const text = await file.text();
      const lines = text.trim().split("\n").filter(Boolean);
      const parsed = [];
      for (let i = 1; i < lines.length; i++) {
        const cols = lines[i].split(/[,;	]/);
        if (cols.length < 3) continue;
        const desc = (cols[1] || "").replace(/"/g, "").trim();
        const amount = Math.abs(parseFloat((cols[2] || "").replace(/[^\d.,-]/g, "").replace(",", ".")));
        if (desc && !isNaN(amount) && amount > 0)
          parsed.push({ date: cols[0] || "", desc, amount });
      }
      if (parsed.length) loadData(parsed);
      else alert("CSV okunamadı.");
    }
  }

  // ── Analiz verileri ─────────────────────────────────────────────────────
  const total = tx.reduce((s, t) => s + t.amount, 0);
  const catMap = {};
  tx.forEach((t) => { catMap[t.cat] = (catMap[t.cat] || 0) + t.amount; });
  const catEntries = Object.entries(catMap).sort((a, b) => b[1] - a[1]);
  const months = [...new Set(tx.map((t) => t.date.slice(0, 7)))].sort();
  const currM = months[months.length - 1] || "";
  const prevM = months.length > 1 ? months[months.length - 2] : null;
  const currTotal = tx.filter((t) => t.date.startsWith(currM)).reduce((s, t) => s + t.amount, 0);
  const prevTotal = prevM ? tx.filter((t) => t.date.startsWith(prevM)).reduce((s, t) => s + t.amount, 0) : null;
  const change = prevTotal ? ((currTotal - prevTotal) / prevTotal) * 100 : 0;
  const MN = ["Oca","Şub","Mar","Nis","May","Haz","Tem","Ağu","Eyl","Eki","Kas","Ara"];

  // ── AI çağrısı ───────────────────────────────────────────────────────────
  async function callAI(prompt) {
    if (!apiKey) { alert("Lütfen Gemini API key'inizi girin."); return; }
    setAiLoading(true);
    setAiText("");
    try {
      const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }]
          }),
        }
      );
      const data = await res.json();
      if (!res.ok) {
        setAiText(`Hata (${res.status}): ${data.error?.message || JSON.stringify(data)}`);
      } else {
        setAiText(data.candidates?.[0]?.content?.parts?.[0]?.text || "Yanıt alınamadı.");
      }
    } catch (e) {
      setAiText("Bağlantı hatası: " + e.message);
    }
    setAiLoading(false);
  }

  function buildPrompt(type) {
    const catSummary = catEntries.map(([c, v]) => `- ${c}: ${Math.round(v).toLocaleString("tr-TR")} ₺ (%${Math.round((v / total) * 100)})`).join("\n");
    if (type === "genel") return `Sen bir kişisel finans danışmanısın. Aşağıdaki harcama verilerini analiz et.\n\nToplam: ${Math.round(total).toLocaleString("tr-TR")} ₺ (${months.length} ay)\nAylık ort: ${Math.round(total / Math.max(months.length, 1)).toLocaleString("tr-TR")} ₺\n\n${catSummary}\n\n1. En dikkat çekici 2-3 harcama kalemi\n2. Somut 3 tasarruf önerisi\n3. Bu profile özel bir finansal hedef\nTürkçe, 250-300 kelime.`;
    if (type === "ipucu") return `Kişisel finans koçusun. Bu harcama dağılımına göre 5 pratik tasarruf ipucu ver. Her biri 1-2 cümle. Türkçe.\n\n${catSummary}`;
    return `Finansal planlama uzmanısın. Aylık ortalama harcama ${Math.round(total / Math.max(months.length, 1)).toLocaleString("tr-TR")} ₺. Kategoriler: ${Object.keys(catMap).join(", ")}. 3 aylık gerçekçi bütçe hedefi öner. Türkçe, kısa ve net.`;
  }

  // ── Render ───────────────────────────────────────────────────────────────
  return (
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "24px 20px", fontFamily: "system-ui, sans-serif" }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
        <div style={{ width: 40, height: 40, borderRadius: 10, background: "#3266ad", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="1.8"><path d="M3 3h18v4H3zM3 10h11v4H3zM3 17h7v4H3z"/><polyline points="17 14 20 17 23 14"/><line x1="20" y1="10" x2="20" y2="17"/></svg>
        </div>
        <div>
          <div style={{ fontSize: 20, fontWeight: 500 }}>Kişisel finans asistanı</div>
          <div style={{ fontSize: 13, color: "#666" }}>Kategorileri kendin belirle, harcamalarını analiz et</div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", borderBottom: "1px solid #e5e5e5", marginBottom: 24 }}>
        {["upload","categories","analysis","ai"].map((t) => (
          <button key={t} onClick={() => setTab(t)} style={{ padding: "10px 18px", fontSize: 13, background: "none", border: "none", borderBottom: tab === t ? "2px solid #3266ad" : "2px solid transparent", color: tab === t ? "#111" : "#888", fontWeight: tab === t ? 500 : 400, cursor: "pointer" }}>
            {{ upload: "Veri yükle", categories: "Kategoriler", analysis: "Analiz", ai: "AI tavsiyeleri" }[t]}
          </button>
        ))}
      </div>

      {/* ── Veri yükle ── */}
      {tab === "upload" && (
        <div>
          <div
            onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
            onDragLeave={() => setDrag(false)}
            onDrop={(e) => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]); }}
            onClick={() => fileRef.current.click()}
            style={{ border: `2px dashed ${drag ? "#3266ad" : "#ccc"}`, borderRadius: 12, padding: "48px 24px", textAlign: "center", cursor: "pointer", background: drag ? "#f0f6ff" : "#fafafa" }}
          >
            <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="#aaa" strokeWidth="1.5" style={{ margin: "0 auto 12px", display: "block" }}><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            <div style={{ fontSize: 15, color: "#555", marginBottom: 4 }}>Excel veya CSV ekstreyi buraya sürükleyin</div>
            <div style={{ fontSize: 12, color: "#aaa" }}>.xlsx · .xls · .csv — Sütun sırası: Tarih, Açıklama, Tutar</div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files[0])} />
          </div>
          <div style={{ textAlign: "center", marginTop: 12 }}>
            <button onClick={() => loadData(DEMO)} style={{ padding: "7px 16px", border: "1px solid #ddd", borderRadius: 8, fontSize: 13, color: "#666", cursor: "pointer", background: "#fff" }}>
              Demo veri ile dene
            </button>
          </div>
        </div>
      )}

      {/* ── Kategoriler ── */}
      {tab === "categories" && (
        <div>
          <p style={{ fontSize: 13, color: "#666", marginBottom: 16 }}>Her kategori için isim, renk ve anahtar kelimeler belirle. Enter ile kelime ekle.</p>
          {cats.map((cat, idx) => (
            <div key={cat.id} style={{ display: "grid", gridTemplateColumns: "32px 1fr 1fr 36px", gap: 8, alignItems: "center", padding: "10px 12px", border: "1px solid #eee", borderRadius: 8, marginBottom: 8, background: "#fff" }}>
              <input type="color" value={cat.color} onChange={(e) => { const c = [...cats]; c[idx].color = e.target.value; setCats(c); }} style={{ width: 28, height: 28, border: "none", borderRadius: 4, cursor: "pointer", padding: 0 }} />
              <input value={cat.name} onChange={(e) => { const c = [...cats]; c[idx].name = e.target.value; setCats(c); }} placeholder="Kategori adı" style={{ border: "1px solid #e0e0e0", borderRadius: 6, padding: "5px 8px", fontSize: 13 }} />
              <div style={{ display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center", border: "1px solid #e0e0e0", borderRadius: 6, padding: "4px 6px", minHeight: 32 }}>
                {cat.keywords.map((kw) => (
                  <span key={kw} style={{ display: "inline-flex", alignItems: "center", gap: 3, background: "#f5f5f5", border: "1px solid #e0e0e0", borderRadius: 3, padding: "1px 6px", fontSize: 11 }}>
                    {kw}
                    <button onClick={() => { const c = [...cats]; c[idx].keywords = c[idx].keywords.filter((k) => k !== kw); setCats(c); }} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 13, color: "#aaa", padding: 0, lineHeight: 1 }}>×</button>
                  </span>
                ))}
                <input placeholder="kelime ekle..." style={{ border: "none", outline: "none", fontSize: 12, minWidth: 80, padding: "1px 2px" }}
                  onKeyDown={(e) => {
                    if (["Enter", ",", " "].includes(e.key)) {
                      e.preventDefault();
                      const val = e.target.value.trim().toLowerCase();
                      if (val && !cat.keywords.includes(val)) { const c = [...cats]; c[idx].keywords.push(val); setCats(c); }
                      e.target.value = "";
                    }
                  }}
                />
              </div>
              {cat.name !== "Diğer" ? (
                <button onClick={() => setCats(cats.filter((_, i) => i !== idx))} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 20, color: "#ccc" }}>×</button>
              ) : <span />}
            </div>
          ))}
          <button onClick={() => setCats([...cats.slice(0, -1), { id: Date.now(), name: "", color: "#3266ad", keywords: [] }, cats[cats.length - 1]])}
            style={{ width: "100%", padding: "9px", border: "1px dashed #ccc", borderRadius: 8, fontSize: 13, color: "#888", cursor: "pointer", background: "none", marginBottom: 12 }}>
            + Yeni kategori ekle
          </button>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
            <button onClick={() => setCats(JSON.parse(JSON.stringify(DEFAULT_CATS)))} style={{ padding: "8px 16px", border: "1px solid #ddd", borderRadius: 8, fontSize: 13, cursor: "pointer", background: "#fff" }}>Sıfırla</button>
            <button onClick={() => setTab("analysis")} style={{ padding: "8px 20px", background: "#3266ad", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, cursor: "pointer" }}>Kaydet ve analizi gör</button>
          </div>
        </div>
      )}

      {/* ── Analiz ── */}
      {tab === "analysis" && (
        !tx.length ? (
          <div style={{ textAlign: "center", padding: "48px", color: "#aaa" }}>Önce veri yükleyin.</div>
        ) : (
          <div>
            {/* Metrikler */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 20 }}>
              {[
                { label: "Toplam harcama", value: `${Math.round(total).toLocaleString("tr-TR")} ₺`, sub: `${months.length} ay` },
                { label: "Bu ay", value: `${Math.round(currTotal).toLocaleString("tr-TR")} ₺`, sub: prevTotal ? `${change > 0 ? "▲" : "▼"} %${Math.abs(change).toFixed(1)} geçen aya göre` : "—", subColor: change > 0 ? "#c0392b" : "#1a7a4a" },
                { label: "Aylık ortalama", value: `${Math.round(total / Math.max(months.length, 1)).toLocaleString("tr-TR")} ₺`, sub: "aylık" },
                { label: "En yüksek", value: catEntries[0]?.[0] || "—", sub: `${Math.round(catEntries[0]?.[1] || 0).toLocaleString("tr-TR")} ₺`, valueSize: 14 },
              ].map((m) => (
                <div key={m.label} style={{ background: "#f5f5f5", borderRadius: 8, padding: "14px 16px" }}>
                  <div style={{ fontSize: 11, color: "#999", marginBottom: 5 }}>{m.label}</div>
                  <div style={{ fontSize: m.valueSize || 20, fontWeight: 500, lineHeight: 1.3 }}>{m.value}</div>
                  <div style={{ fontSize: 11, color: m.subColor || "#aaa", marginTop: 2 }}>{m.sub}</div>
                </div>
              ))}
            </div>

            {/* Grafikler */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 20 }}>
              <div style={{ border: "1px solid #eee", borderRadius: 12, padding: "16px" }}>
                <div style={{ fontSize: 12, color: "#aaa", fontWeight: 500, marginBottom: 10, textTransform: "uppercase", letterSpacing: ".04em" }}>Kategori dağılımı</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>
                  {catEntries.map(([n, v]) => (
                    <span key={n} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#555" }}>
                      <span style={{ width: 8, height: 8, borderRadius: 2, background: getColor(n, cats), display: "inline-block" }} />
                      {n} {Math.round((v / total) * 100)}%
                    </span>
                  ))}
                </div>
                <div style={{ height: 200 }}>
                  <Doughnut
                    data={{ labels: catEntries.map(([n]) => n), datasets: [{ data: catEntries.map(([, v]) => Math.round(v)), backgroundColor: catEntries.map(([n]) => getColor(n, cats)), borderWidth: 2, borderColor: "#fff" }] }}
                    options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }}
                  />
                </div>
              </div>
              <div style={{ border: "1px solid #eee", borderRadius: 12, padding: "16px" }}>
                <div style={{ fontSize: 12, color: "#aaa", fontWeight: 500, marginBottom: 10, textTransform: "uppercase", letterSpacing: ".04em" }}>Aylık trend</div>
                <div style={{ height: 232 }}>
                  <Bar
                    data={{ labels: months.map((m) => { const [y, mo] = m.split("-"); return MN[parseInt(mo) - 1] + " '" + y.slice(2); }), datasets: [{ data: months.map((m) => Math.round(tx.filter((t) => t.date.startsWith(m)).reduce((s, t) => s + t.amount, 0))), backgroundColor: "#3266ad", borderRadius: 4 }] }}
                    options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { ticks: { callback: (v) => v.toLocaleString("tr-TR") + "₺" } }, x: { ticks: { autoSkip: false } } } }}
                  />
                </div>
              </div>
            </div>

            {/* İşlemler tablosu */}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontSize: 13, fontWeight: 500, color: "#555" }}>Son işlemler</div>
              <select value={filter} onChange={(e) => setFilter(e.target.value)} style={{ fontSize: 12, padding: "4px 8px", border: "1px solid #e0e0e0", borderRadius: 6 }}>
                <option value="">Tüm kategoriler</option>
                {catEntries.map(([n]) => <option key={n} value={n}>{n}</option>)}
              </select>
            </div>
            <div style={{ border: "1px solid #eee", borderRadius: 12, overflow: "hidden" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: "#f9f9f9" }}>
                    {["Tarih", "Açıklama", "Kategori", "Tutar"].map((h) => (
                      <th key={h} style={{ padding: "9px 14px", textAlign: h === "Tutar" ? "right" : "left", fontWeight: 500, fontSize: 11, color: "#aaa", textTransform: "uppercase", letterSpacing: ".04em", borderBottom: "1px solid #eee" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[...tx].filter((t) => !filter || t.cat === filter).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 20).map((t, i) => {
                    const col = getColor(t.cat, cats);
                    return (
                      <tr key={i} style={{ borderBottom: "1px solid #f5f5f5" }}>
                        <td style={{ padding: "9px 14px", color: "#777" }}>{t.date}</td>
                        <td style={{ padding: "9px 14px" }}>{t.desc}</td>
                        <td style={{ padding: "9px 14px" }}><span style={{ background: col + "22", color: col, padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 500 }}>{t.cat}</span></td>
                        <td style={{ padding: "9px 14px", textAlign: "right", fontWeight: 500 }}>{t.amount.toLocaleString("tr-TR")} ₺</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )
      )}

      {/* ── AI Tavsiyeleri ── */}
      {tab === "ai" && (
        !tx.length ? (
          <div style={{ textAlign: "center", padding: "48px", color: "#aaa" }}>Önce veri yükleyin.</div>
        ) : (
          <div>
            {/* AI sonucu */}
            <div style={{ border: "1px solid #eee", borderRadius: 12, padding: "20px", marginBottom: 16, minHeight: 80 }}>
              {aiLoading ? (
                <div style={{ color: "#aaa", fontSize: 13 }}>Analiz yapılıyor...</div>
              ) : aiText ? (
                <div style={{ fontSize: 14, lineHeight: 1.75, whiteSpace: "pre-wrap", color: "#333" }}>{aiText}</div>
              ) : (
                <div style={{ color: "#bbb", fontSize: 13 }}>Aşağıdaki butonlardan birini seçin.</div>
              )}
            </div>

            {/* Butonlar */}
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button onClick={() => callAI(buildPrompt("genel"))} disabled={aiLoading} style={{ padding: "9px 20px", background: "#3266ad", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                Genel analiz
              </button>
              <button onClick={() => callAI(buildPrompt("ipucu"))} disabled={aiLoading} style={{ padding: "9px 16px", border: "1px solid #ddd", borderRadius: 8, fontSize: 13, cursor: "pointer", background: "#fff" }}>
                Tasarruf ipuçları
              </button>
              <button onClick={() => callAI(buildPrompt("hedef"))} disabled={aiLoading} style={{ padding: "9px 16px", border: "1px solid #ddd", borderRadius: 8, fontSize: 13, cursor: "pointer", background: "#fff" }}>
                Bütçe hedefi öner
              </button>
            </div>
          </div>
        )
      )}
    </div>
  );
}
