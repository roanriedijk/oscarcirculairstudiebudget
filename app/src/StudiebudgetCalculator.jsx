import React, { useState, useMemo, useRef } from "react";
import { Plus, Trash2, Calculator, RotateCcw, AlertTriangle, Download, Upload, Loader2 } from "lucide-react";
import ExcelJS from "exceljs";

const RATE = 0.02;
const LOGO = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPAAAACwCAYAAAAxDeP+AAA6UklEQVR4nO2deXxcVfn/P89z7r2TdC+lCy2bgCzFtqTTsmuRtQJfRZFFURRE0qZJs7R8Qb9qjIoKXZK0WRpcavGnYnFDUFsKYlUEaSctWxFEkK2lC23pkmTuvec8vz9mpqQlSedOJpmmue/XK+3rNXPvuc89cz5zH+qlN1/buK5sLxcyj40QjbRraqLHDrbtPzgO32FEBriuyYrtIoluafIHGmxb/Fkh89fG2uh99TUFE7uT9vLl16gl1dEKi9XjlkWfAcR2XZNWTyFdJNFlhe8LlKLzlOIVDTXReU2V0QFB06q9O3piBEP/aDvquyAang1bRZDIW8Ewx1bVS2qjy+rqxg/qVqK9QN2C6BzbpkUiiPh+z4UUS0vAjTUFxQx6zLJ4mueZfa1PT2GMIO4aENExlqKfNNRMXtbwvQnDM02vbl7BcXmEP1g2T4tnSbgdsa+wAZZt8XWK+O8NCyd/Z8GCs/ODplVTc9modqe1FZP+TYmqcrN2J5Kv2jKJdj7dq3DOrXJq+ZzDqCxPQUQhLGFkBFuNEGimQqNPZTAZqE/KqhJIMZ";

const fmtEUR = (n) =>
  new Intl.NumberFormat("nl-NL", {
    style: "currency",
    currency: "EUR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(isFinite(n) ? n : 0);

const fmtDate = (d) =>
  d.toLocaleDateString("nl-NL", { day: "numeric", month: "short", year: "numeric" });

const TODAY = (() => {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
})();

const windowStartFor = (endDate) => {
  const d = new Date(endDate);
  d.setFullYear(d.getFullYear() - 2);
  return d;
};

const toDate = (s) => {
  if (!s) return null;
  const d = new Date(s + "T00:00:00");
  return isNaN(d) ? null : d;
};

const inputISO = (d) => d.toISOString().slice(0, 10);
const daysInMonth = (year, month) => new Date(year, month + 1, 0).getDate();

const grossForRange = (start, end, monthly) => {
  if (!start || !end || end < start) return 0;
  let total = 0;
  let cur = new Date(start.getFullYear(), start.getMonth(), 1);
  while (cur <= end) {
    const y = cur.getFullYear();
    const m = cur.getMonth();
    const dim = daysInMonth(y, m);
    const monthStart = new Date(y, m, 1);
    const monthEnd = new Date(y, m, dim);
    const overlapStart = monthStart > start ? monthStart : start;
    const overlapEnd = monthEnd < end ? monthEnd : end;
    if (overlapEnd >= overlapStart) {
      const days = Math.round((overlapEnd - overlapStart) / 86400000) + 1;
      total += monthly * (days / dim);
    }
    cur = new Date(y, m + 1, 1);
  }
  return total;
};

const clampTo = (start, end, winStart, winEnd) => {
  const s = start > winStart ? start : winStart;
  const e = end < winEnd ? end : winEnd;
  if (e < s) return null;
  return { start: s, end: e };
};

const upTo = (start, end, cap) => {
  if (!start || !end) return null;
  const e = end < cap ? end : cap;
  if (e < start) return null;
  return { start, end: e };
};

let _id = 1;
const newPeriod = () => ({ id: _id++, start: "", end: "", monthly: "" });

export default function StudiebudgetCalculator() {
  const [name, setName] = useState("");
  const [periods, setPeriods] = useState([newPeriod()]);
  const [alreadyUsed, setAlreadyUsed] = useState("");
  const [extracting, setExtracting] = useState(false);
  const [extractError, setExtractError] = useState(null);
  const [calculated, setCalculated] = useState(false);
  const fileInputRef = useRef(null);

  const handleScreenshot = async (file) => {
    if (!file) return;
    setExtractError(null);
    setExtracting(true);
    try {
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("Kon de afbeelding niet lezen"));
        r.readAsDataURL(file);
      });
      const mediaType = file.type || "image/png";

      const prompt =
        'Dit is een schermafbeelding van een Nederlandse salaristabel met kolommen zoals "Ingangsdatum", "Type werkschema", "Contracturen" en "Bruto salaris". ' +
        'Haal voor elke rij de ingangsdatum en het bruto maandsalaris (het eurobedrag) op. ' +
        'Geef UITSLUITEND geldige JSON terug, zonder uitleg of markdown, in dit formaat: ' +
        '{"rows":[{"start":"YYYY-MM-DD","monthly":1340.00}]}. ' +
        'De datum moet als ISO-datum (YYYY-MM-DD). Het bedrag als getal zonder valutateken of duizendscheiding (gebruik een punt als decimaalteken). ' +
        'Negeer eventuele percentages tussen haakjes; neem alleen het eurobedrag.';

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          messages: [
            {
              role: "user",
              content: [
                { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
                { type: "text", text: prompt },
              ],
            },
          ],
        }),
      });

      const data = await response.json();
      const text = (data.content || [])
        .map((b) => (b.type === "text" ? b.text : ""))
        .join("")
        .replace(/```json|```/g, "")
        .trim();
      const parsed = JSON.parse(text);
      const raw = (parsed.rows || [])
        .filter((r) => r.start)
        .map((r) => ({ start: r.start, monthly: r.monthly }));

      if (raw.length === 0) {
        setExtractError("Geen regels gevonden in de afbeelding. Controleer de schermafbeelding of vul handmatig in.");
        setExtracting(false);
        return;
      }

      raw.sort((a, b) => new Date(a.start) - new Date(b.start));

      const built = raw.map((r, i) => {
        let end = "";
        if (i < raw.length - 1) {
          const next = new Date(raw[i + 1].start + "T00:00:00");
          next.setDate(next.getDate() - 1);
          end = inputISO(next);
        }
        return {
          id: _id++,
          start: r.start,
          end,
          monthly: r.monthly != null ? String(r.monthly) : "",
        };
      });

      setPeriods(built);
      setCalculated(false);
      setExtracting(false);
    } catch (e) {
      setExtractError(
        "Het automatisch uitlezen is niet gelukt. Probeer een duidelijkere schermafbeelding of vul de regels handmatig in."
      );
      setExtracting(false);
    }
  };

  const update = (id, field, value) => {
    setCalculated(false);
    setPeriods((p) => p.map((r) => (r.id === id ? { ...r, [field]: value } : r)));
  };
  const addPeriod = () => { setCalculated(false); setPeriods((p) => [...p, newPeriod()]); };
  const removePeriod = (id) => {
    setCalculated(false);
    setPeriods((p) => (p.length > 1 ? p.filter((r) => r.id !== id) : p));
  };
  const reset = () => {
    setName("");
    setPeriods([newPeriod()]);
    setAlreadyUsed("");
    setCalculated(false);
    setExtractError(null);
  };

  const contractEnd = useMemo(() => {
    let latest = null;
    periods.forEach((p) => {
      const e = toDate(p.end);
      if (e && (!latest || e > latest)) latest = e;
    });
    return latest;
  }, [periods]);

  const windowStart = useMemo(
    () => (contractEnd ? windowStartFor(contractEnd) : null),
    [contractEnd]
  );

  const rows = useMemo(
    () =>
      periods.map((p) => {
        const start = toDate(p.start);
        const end = toDate(p.end);
        const monthly = parseFloat(p.monthly) || 0;

        let note = null;
        let eligibleGross = 0;
        let pastGross = 0;
        let futureGross = 0;
        let fullGross = 0;
        let clamped = null;

        if (start && end && end >= start && windowStart) {
          fullGross = grossForRange(start, end, monthly);
          clamped = clampTo(start, end, windowStart, contractEnd);
          if (!clamped) {
            note = "Buiten de periode van 2 jaar — telt niet mee";
          } else {
            eligibleGross = grossForRange(clamped.start, clamped.end, monthly);
            const pastPart = upTo(clamped.start, clamped.end, TODAY);
            if (pastPart) {
              pastGross = grossForRange(pastPart.start, pastPart.end, monthly);
            }
            futureGross = eligibleGross - pastGross;
            if (clamped.start > start) {
              note = `Deels vóór de geldige periode — meegerekend vanaf ${fmtDate(clamped.start)}`;
            }
          }
        } else if (start && end && end < start) {
          note = "Einddatum ligt vóór de startdatum";
        } else if (start && !end && monthly > 0) {
          note = "Vul een einddatum in om deze regel mee te rekenen";
        }

        return {
          ...p,
          startD: start,
          endD: end,
          monthlyNum: monthly,
          fullGross,
          eligibleGross,
          pastGross,
          futureGross,
          budget: eligibleGross * RATE,
          clamped,
          note,
        };
      }),
    [periods, windowStart, contractEnd]
  );

  const totalEligibleGross = rows.reduce((s, r) => s + r.eligibleGross, 0);
  const totalBudget = rows.reduce((s, r) => s + r.budget, 0);
  const accruedBudget = rows.reduce((s, r) => s + r.pastGross * RATE, 0);
  const advanceBudget = rows.reduce((s, r) => s + r.futureGross * RATE, 0);
  const used = parseFloat(alreadyUsed) || 0;
  const remaining = totalBudget - used;
  const overspent = remaining < -0.005;

  const exportExcel = async () => {
    const BLUE_HEX = "FF1B3F8F";
    const BLUE_DEEP_HEX = "FF13306E";
    const GOLD_HEX = "FF908830";
    const LIGHT_HEX = "FFF3F6FC";
    const GREY_HEX = "FF8A93A8";

    const wb = new ExcelJS.Workbook();
    wb.creator = "Oscar Circulair";
    wb.created = new Date();
    const ws = wb.addWorksheet("Studiebudget", {
      views: [{ showGridLines: false }],
      properties: { defaultRowHeight: 18 },
    });

    ws.columns = [
      { width: 18 }, { width: 18 }, { width: 20 }, { width: 34 }, { width: 22 }, { width: 20 },
    ];

    try {
      const imgId = wb.addImage({ base64: LOGO, extension: "png" });
      ws.addImage(imgId, { tl: { col: 0.15, row: 0.2 }, ext: { width: 150, height: 110 } });
    } catch (e) {}
    ws.getRow(1).height = 28;
    ws.getRow(2).height = 28;
    ws.getRow(3).height = 28;
    ws.getRow(4).height = 12;

    ws.mergeCells("C1:F1");
    const tCell = ws.getCell("C1");
    tCell.value = "Studiebudget";
    tCell.font = { name: "Arial", size: 20, bold: true, color: { argb: BLUE_DEEP_HEX } };
    tCell.alignment = { vertical: "middle", horizontal: "right" };

    ws.mergeCells("C2:F2");
    const sCell = ws.getCell("C2");
    sCell.value = "Oscar Circulair · HR";
    sCell.font = { name: "Arial", size: 11, bold: true, color: { argb: GOLD_HEX } };
    sCell.alignment = { vertical: "middle", horizontal: "right" };

    let r = 6;
    const meta = [
      ["Medewerker", name || "—"],
      ["Geldige periode", windowStart && contractEnd ? `${fmtDate(windowStart)} t/m ${fmtDate(contractEnd)}` : "—"],
      ["Peildatum (vandaag)", fmtDate(TODAY)],
    ];
    meta.forEach(([k, v]) => {
      const kc = ws.getCell(`A${r}`);
      kc.value = k;
      kc.font = { name: "Arial", size: 10, bold: true, color: { argb: GREY_HEX } };
      const vc = ws.getCell(`B${r}`);
      vc.value = v;
      vc.font = { name: "Arial", size: 11, color: { argb: "FF1A2233" } };
      ws.mergeCells(`B${r}:D${r}`);
      r++;
    });
    r++;

    const headerRow = r;
    const headers = [
      "Startdatum", "Einddatum", "Bruto maandsalaris",
      "Meegerekende periode", "Meetellend brutoloon", "Studiebudget (2%)",
    ];
    headers.forEach((h, i) => {
      const cell = ws.getCell(headerRow, i + 1);
      cell.value = h;
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE_HEX } };
      cell.alignment = { vertical: "middle", horizontal: i >= 2 ? "right" : "left", wrapText: true };
      cell.border = { bottom: { style: "thin", color: { argb: BLUE_DEEP_HEX } } };
    });
    ws.getRow(headerRow).height = 26;
    r++;

    const euroFmt = "€ #,##0.00";
    const dataRows = rows.filter((row) => row.startD && row.endD && row.endD >= row.startD);
    dataRows.forEach((row, idx) => {
      const vals = [
        fmtDate(row.startD),
        fmtDate(row.endD),
        row.monthlyNum,
        row.clamped ? `${fmtDate(row.clamped.start)} t/m ${fmtDate(row.clamped.end)}` : "Buiten venster",
        Number(row.eligibleGross.toFixed(2)),
        Number(row.budget.toFixed(2)),
      ];
      vals.forEach((v, i) => {
        const cell = ws.getCell(r, i + 1);
        cell.value = v;
        cell.font = { name: "Arial", size: 10, color: { argb: "FF1A2233" } };
        cell.alignment = { vertical: "middle", horizontal: i >= 2 ? "right" : "left" };
        if (i === 2 || i >= 4) cell.numFmt = euroFmt;
        if (idx % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT_HEX } };
        cell.border = { bottom: { style: "hair", color: { argb: "FFD3DBEB" } } };
      });
      r++;
    });
    r++;

    const totals = [
      ["Totaal meetellend brutoloon", Number(totalEligibleGross.toFixed(2)), false],
      ["Totaal studiebudget (2%)", Number(totalBudget.toFixed(2)), true],
      ["   waarvan reeds opgebouwd (t/m vandaag)", Number(accruedBudget.toFixed(2)), false],
      ["   waarvan vooraf op te nemen (toekomst)", Number(advanceBudget.toFixed(2)), false],
    ];
    if (used > 0) {
      totals.push(["Al gebruikt", Number(used.toFixed(2)), false]);
      totals.push(["Resterend", Number(remaining.toFixed(2)), true]);
    }
    totals.forEach(([label, val, highlight]) => {
      ws.mergeCells(`A${r}:D${r}`);
      const lc = ws.getCell(`A${r}`);
      lc.value = label;
      lc.alignment = { horizontal: "right", vertical: "middle" };
      lc.font = { name: "Arial", size: 10, bold: highlight, color: { argb: highlight ? BLUE_DEEP_HEX : "FF1A2233" } };
      const ec = ws.getCell(`E${r}`);
      ec.value = val;
      ec.numFmt = euroFmt;
      ec.alignment = { horizontal: "right", vertical: "middle" };
      ec.font = { name: "Arial", size: highlight ? 12 : 10, bold: highlight, color: { argb: highlight ? BLUE_DEEP_HEX : "FF1A2233" } };
      if (highlight) {
        ["A", "B", "C", "D", "E"].forEach((col) => {
          ws.getCell(`${col}${r}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT_HEX } };
        });
      }
      ws.getRow(r).height = highlight ? 22 : 18;
      r++;
    });

    r += 1;
    ws.mergeCells(`A${r}:F${r}`);
    const noteCell = ws.getCell(`A${r}`);
    noteCell.value =
      "Het bedrag 'vooraf op te nemen' betreft toekomstig contractloon en dient bij eerder vertrek te worden terugbetaald.";
    noteCell.font = { name: "Arial", size: 9, italic: true, color: { argb: GREY_HEX } };
    noteCell.alignment = { wrapText: true, vertical: "top" };
    ws.getRow(r).height = 28;

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const safe = (name || "medewerker").replace(/[^\w\-]+/g, "_");
    a.download = `studiebudget_${safe}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const hasResult = rows.some((r) => r.budget > 0);

  return (
    <div style={wrap}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } } .spin { animation: spin 0.8s linear infinite; }`}</style>
      <div style={{ maxWidth: 780, margin: "0 auto" }}>
        <div style={{ marginBottom: 28 }}>
          <div style={brandRow}>
            <img src={LOGO} alt="Oscar Circulair" style={{ height: 88, width: "auto" }} />
          </div>
          <div style={eyebrow}>HR · Studiebudget</div>
          <h1 style={h1}>Studiebudget-calculator</h1>
          <p style={lede}>
            Voer per periode het <strong>bruto maandsalaris</strong> in. De app
            rekent zelf uit wat het totale brutoloon over die periode is en
            berekent daarvan 2% als studiebudget. Het budget mag besteed worden
            over een doorlopende periode van twee jaar, gerekend vanaf de{" "}
            <strong>einddatum van het contract</strong> (de laatste einddatum die
            je invult). Loon van vóór dat venster vervalt; toekomstig loon binnen
            het lopende contract telt wél mee.
            {windowStart && contractEnd && (
              <>
                {" "}Voor de ingevulde regels loopt het venster van{" "}
                <strong style={{ color: BLUE_DEEP }}>{fmtDate(windowStart)}</strong> t/m{" "}
                <strong style={{ color: BLUE_DEEP }}>{fmtDate(contractEnd)}</strong>.
              </>
            )}
          </p>
        </div>

        <div style={card}>
          <div style={uploadZone}>
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, color: BLUE_DEEP, fontSize: 15, marginBottom: 4 }}>
                Vul automatisch in vanuit een schermafbeelding
              </div>
              <div style={{ fontSize: 13, color: "#4a5570", lineHeight: 1.45 }}>
                Upload een schermafbeelding van de salaristabel. De app leest de
                ingangsdatums en bruto maandsalarissen uit en vult de regels in.
                Daarna controleer je en klik je op Uitrekenen.
              </div>
              {extractError && (
                <div style={{ marginTop: 8, fontSize: 12.5, color: "#c0392b", display: "flex", alignItems: "center", gap: 6 }}>
                  <AlertTriangle size={13} /> {extractError}
                </div>
              )}
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*"
              style={{ display: "none" }}
              onChange={(e) => handleScreenshot(e.target.files && e.target.files[0])}
            />
            <button
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              disabled={extracting}
              style={{ ...uploadBtn, opacity: extracting ? 0.7 : 1, cursor: extracting ? "wait" : "pointer" }}
            >
              {extracting ? <Loader2 size={16} className="spin" /> : <Upload size={16} />}
              {extracting ? "Bezig met uitlezen…" : "Schermafbeelding uploaden"}
            </button>
          </div>

          <label style={{ display: "block", marginBottom: 22 }}>
            <span style={fieldLabel}>Naam medewerker (optioneel)</span>
            <input
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="bijv. Sanne de Vries"
              style={input}
            />
          </label>

          <div style={headerGrid}>
            <span>Startdatum</span>
            <span>Einddatum</span>
            <span>Bruto maandsalaris (€)</span>
            <span style={{ width: 36 }} />
          </div>

          {rows.map((r) => (
            <div key={r.id} style={{ marginBottom: 12 }}>
              <div style={rowGrid}>
                <input
                  type="date"
                  value={r.start}
                  onChange={(e) => update(r.id, "start", e.target.value)}
                  style={input}
                />
                <input
                  type="date"
                  value={r.end}
                  onChange={(e) => update(r.id, "end", e.target.value)}
                  style={input}
                />
                <input
                  type="number"
                  inputMode="decimal"
                  min="0"
                  value={r.monthly}
                  onChange={(e) => update(r.id, "monthly", e.target.value)}
                  placeholder="0,00"
                  style={input}
                />
                <button
                  onClick={() => removePeriod(r.id)}
                  disabled={periods.length === 1}
                  title="Periode verwijderen"
                  style={{
                    ...iconBtn,
                    opacity: periods.length === 1 ? 0.35 : 1,
                    cursor: periods.length === 1 ? "not-allowed" : "pointer",
                  }}
                >
                  <Trash2 size={16} />
                </button>
              </div>
              {(r.note || r.budget > 0) && (
                <div style={{ marginTop: 5, display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                  {r.note ? (
                    <span
                      style={{
                        fontSize: 12.5,
                        color: r.note.includes("Buiten") || r.note.includes("vóór") ? "#c0392b" : "#b9770e",
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                      }}
                    >
                      <AlertTriangle size={13} /> {r.note}
                    </span>
                  ) : <span />}
                  {r.budget > 0 && (
                    <span style={{ fontSize: 12.5, color: "#4a5570" }}>
                      Brutoloon periode: <strong>{fmtEUR(r.eligibleGross)}</strong> → 2% = <strong>{fmtEUR(r.budget)}</strong>
                    </span>
                  )}
                </div>
              )}
            </div>
          ))}

          <button onClick={addPeriod} style={addBtn}>
            <Plus size={16} /> Periode toevoegen
          </button>

          <label style={{ display: "block", marginTop: 24 }}>
            <span style={fieldLabel}>Al gebruikt (€, optioneel)</span>
            <input
              type="number"
              inputMode="decimal"
              min="0"
              value={alreadyUsed}
              onChange={(e) => { setCalculated(false); setAlreadyUsed(e.target.value); }}
              placeholder="0,00"
              style={{ ...input, maxWidth: 220 }}
            />
          </label>

          <button
            onClick={() => setCalculated(true)}
            disabled={!hasResult}
            style={{
              ...calcBtn,
              opacity: hasResult ? 1 : 0.45,
              cursor: hasResult ? "pointer" : "not-allowed",
            }}
          >
            <Calculator size={17} /> Uitrekenen
          </button>
        </div>

        {calculated && (
          <div style={resultCard}>
            <div style={resultHead}>
              <Calculator size={16} />
              {name ? `Resultaat voor ${name}` : "Resultaat"}
            </div>

            <div style={{ display: "flex", flexWrap: "wrap", gap: 28 }}>
              <div>
                <div style={statLabel}>Totaal brutoloon periode</div>
                <div style={statValue}>{fmtEUR(totalEligibleGross)}</div>
              </div>
              <div>
                <div style={statLabel}>Totaal budget (2%)</div>
                <div style={statValue}>{fmtEUR(totalBudget)}</div>
              </div>
              {used > 0 && (
                <div>
                  <div style={statLabel}>Al gebruikt</div>
                  <div style={statValue}>−{fmtEUR(used)}</div>
                </div>
              )}
              <div>
                <div style={statLabel}>{used > 0 ? "Resterend" : "Beschikbaar"}</div>
                <div style={{ ...statValue, fontSize: 34, color: overspent ? "#ffd9d2" : "#fff" }}>
                  {fmtEUR(remaining)}
                </div>
              </div>
            </div>

            {advanceBudget > 0.005 && (
              <div style={splitPanel}>
                <div style={splitItem}>
                  <div style={splitLabel}>Reeds opgebouwd t/m {fmtDate(TODAY)}</div>
                  <div style={splitValue}>{fmtEUR(accruedBudget)}</div>
                </div>
                <div style={splitDivider} />
                <div style={splitItem}>
                  <div style={splitLabel}>Vooraf op te nemen (toekomstig contract)</div>
                  <div style={splitValue}>{fmtEUR(advanceBudget)}</div>
                  <div style={splitHint}>Terug te betalen bij eerder vertrek</div>
                </div>
              </div>
            )}

            {overspent && (
              <div style={overspentNote}>
                <AlertTriangle size={14} /> Deze medewerker heeft{" "}
                {fmtEUR(used - totalBudget)} méér gebruikt dan het beschikbare budget.
              </div>
            )}

            {hasResult && (
              <div style={breakdown}>
                {rows
                  .filter((r) => r.budget > 0)
                  .map((r) => (
                    <div key={r.id} style={breakdownRow}>
                      <span>
                        {fmtDate(r.clamped.start)} – {fmtDate(r.clamped.end)} · {fmtEUR(r.monthlyNum)}/mnd
                      </span>
                      <span>
                        {fmtEUR(r.eligibleGross)} × 2% = {fmtEUR(r.budget)}
                      </span>
                    </div>
                  ))}
              </div>
            )}
          </div>
        )}

        {calculated && (
          <div style={{ display: "flex", gap: 14, alignItems: "center", marginTop: 16 }}>
            <button
              onClick={exportExcel}
              disabled={!hasResult}
              style={{ ...exportBtn, opacity: hasResult ? 1 : 0.45, cursor: hasResult ? "pointer" : "not-allowed" }}
            >
              <Download size={16} /> Exporteer naar Excel
            </button>
          </div>
        )}

        <div style={{ marginTop: 16 }}>
          <button onClick={reset} style={resetBtn}>
            <RotateCcw size={14} /> Opnieuw
          </button>
        </div>
      </div>
    </div>
  );
}

const BLUE = "#1B3F8F";
const GOLD = "#908830";
const BLUE_DEEP = "#13306e";

const wrap = {
  minHeight: "100vh",
  background: "#f4f6fb",
  fontFamily: "'Inter', system-ui, sans-serif",
  color: "#1a2233",
  padding: "32px 16px",
};
const brandRow = { display: "flex", alignItems: "center", gap: 10, marginBottom: 18 };
const eyebrow = {
  fontSize: 12,
  letterSpacing: "0.18em",
  textTransform: "uppercase",
  color: GOLD,
  fontWeight: 700,
  marginBottom: 8,
};
const h1 = {
  fontSize: 32,
  fontWeight: 800,
  margin: 0,
  lineHeight: 1.1,
  letterSpacing: "-0.02em",
  color: BLUE_DEEP,
};
const lede = { color: "#4a5570", marginTop: 10, fontSize: 15, lineHeight: 1.55 };
const card = {
  background: "#fff",
  borderRadius: 16,
  padding: 24,
  border: "1px solid #e2e6f0",
  boxShadow: "0 1px 3px rgba(20,40,80,0.06)",
};
const uploadZone = {
  display: "flex",
  gap: 16,
  alignItems: "center",
  flexWrap: "wrap",
  background: "#f3f6fc",
  border: `1px dashed ${BLUE}`,
  borderRadius: 12,
  padding: "16px 18px",
  marginBottom: 24,
};
const uploadBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 8,
  background: BLUE,
  border: "none",
  color: "#fff",
  borderRadius: 9,
  padding: "11px 18px",
  fontSize: 14.5,
  fontWeight: 700,
  whiteSpace: "nowrap",
};
const fieldLabel = { fontSize: 13, color: "#4a5570", display: "block", marginBottom: 6 };
const headerGrid = {
  display: "grid",
  gridTemplateColumns: "1fr 1fr 1fr auto",
  gap: 12,
  fontSize: 12,
  color: "#8a93a8",
  textTransform: "uppercase",
  letterSpacing: "0.08em",
  marginBottom: 8,
  paddingRight: 4,
};
const rowGrid = {
  display: "grid",
  gridTemplateColumns: "1fr 1fr 1fr auto",
  gap: 12,
  alignItems: "center",
};
const input = {
  width: "100%",
  boxSizing: "border-box",
  background: "#f7f9fd",
  border: "1px solid #d3dbeb",
  borderRadius: 9,
  padding: "10px 12px",
  color: "#1a2233",
  fontSize: 15,
  outline: "none",
};
const iconBtn = {
  width: 36,
  height: 36,
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  background: "#f7f9fd",
  border: "1px solid #d3dbeb",
  borderRadius: 9,
  color: "#8a93a8",
};
const addBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  marginTop: 6,
  background: "transparent",
  border: `1px dashed ${BLUE}`,
  color: BLUE,
  borderRadius: 9,
  padding: "9px 14px",
  fontSize: 14,
  cursor: "pointer",
  fontWeight: 600,
};
const resultCard = {
  background: `linear-gradient(135deg, ${BLUE}, #14306f)`,
  borderRadius: 16,
  padding: 24,
  marginTop: 18,
};
const resultHead = {
  display: "flex",
  alignItems: "center",
  gap: 8,
  fontSize: 13,
  color: "#dfe7fb",
  marginBottom: 14,
  fontWeight: 600,
};
const statLabel = {
  fontSize: 12,
  color: "#cdd9f5",
  textTransform: "uppercase",
  letterSpacing: "0.08em",
  marginBottom: 4,
};
const statValue = { fontSize: 24, fontWeight: 800, letterSpacing: "-0.02em", color: "#fff" };
const splitPanel = {
  marginTop: 18,
  display: "flex",
  alignItems: "stretch",
  gap: 18,
  background: "rgba(255,255,255,0.12)",
  borderRadius: 12,
  padding: "14px 18px",
  flexWrap: "wrap",
};
const splitItem = { flex: 1, minWidth: 180 };
const splitDivider = { width: 1, background: "rgba(255,255,255,0.25)" };
const splitLabel = {
  fontSize: 11.5,
  color: "#cdd9f5",
  textTransform: "uppercase",
  letterSpacing: "0.07em",
  marginBottom: 4,
};
const splitValue = { fontSize: 20, fontWeight: 800, color: "#fff" };
const splitHint = { fontSize: 11.5, color: "#f0e6b8", marginTop: 3, fontStyle: "italic" };
const overspentNote = {
  marginTop: 16,
  background: "rgba(255,255,255,0.92)",
  border: "1px solid #f0b8b0",
  borderRadius: 10,
  padding: "10px 14px",
  fontSize: 13.5,
  color: "#7a1f15",
  display: "flex",
  alignItems: "center",
  gap: 8,
  fontWeight: 600,
};
const breakdown = {
  marginTop: 18,
  paddingTop: 16,
  borderTop: "1px solid rgba(255,255,255,0.3)",
  fontSize: 13,
  color: "#dfe7fb",
};
const breakdownRow = {
  display: "flex",
  justifyContent: "space-between",
  gap: 12,
  padding: "3px 0",
};
const exportBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 8,
  background: GOLD,
  border: "none",
  color: "#fff",
  borderRadius: 9,
  padding: "11px 18px",
  fontSize: 14.5,
  fontWeight: 700,
};
const calcBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 9,
  marginTop: 22,
  background: BLUE,
  border: "none",
  color: "#fff",
  borderRadius: 10,
  padding: "13px 26px",
  fontSize: 16,
  fontWeight: 800,
  letterSpacing: "-0.01em",
};
const resetBtn = {
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  background: "transparent",
  border: "none",
  color: "#8a93a8",
  fontSize: 13,
  cursor: "pointer",
};
