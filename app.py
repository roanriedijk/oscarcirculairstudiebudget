
import React, { useState, useMemo, useRef } from "react";
import { Plus, Trash2, Calculator, RotateCcw, AlertTriangle, Download, Upload, Loader2 } from "lucide-react";
import ExcelJS from "exceljs";

const RATE = 0.02;
const LOGO = "data:image/png;base64,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";

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

// The spending window runs 2 years back from the contract end date (the latest
// end date entered). Given that end date, return the window start.
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

// Given a clamped date range and a monthly salary, sum the gross by
// pro-rating each calendar month by the share of its days that fall in range.
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

// Clamp a range to [winStart, winEnd]. Returns {start,end} or null if no overlap.
const clampTo = (start, end, winStart, winEnd) => {
  const s = start > winStart ? start : winStart;
  const e = end < winEnd ? end : winEnd;
  if (e < s) return null;
  return { start: s, end: e };
};

// Intersect a range with [, cap] (everything up to and including cap).
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

  // Read an uploaded screenshot, send to the Anthropic API for table extraction,
  // then fill the period rows. Newest row is on top in the source; we sort
  // oldest-first and derive each end date as the day before the next start.
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
        .replace(/``json|``/g, "")
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

      // Sort oldest first by start date.
      raw.sort((a, b) => new Date(a.start) - new Date(b.start));

      // Derive end date = day before next row's start. Last (newest) row: leave empty.
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

  // The contract end date = latest end date entered across all rows.
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
          // Clamp to the spending window: 2 years back from contract end,
          // up to the contract end itself (future contract months count).
          clamped = clampTo(start, end, windowStart, contractEnd);
          if (!clamped) {
            note = "Buiten de periode van 2 jaar — telt niet mee";
          } else {
            eligibleGross = grossForRange(clamped.start, clamped.end, monthly);
            // Split into already accrued (up to today) and taken in advance (after today).
            const pastPart = upTo(clamped.start, clamped.end, TODAY);
            if (pastPart) {
              pastGross = grossForRange(pastPart.start, pastPart.end, monthly);
            }
            futureGross = eligibleGross - pastGross;
            if (clamped.start > start) {
              note = `Deels vóór de geldige periode — meegerekend vanaf ${fmtDate(
                clamped.start
              )}`;
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

    // Column widths (A..F)
    ws.columns = [
      { width: 18 }, { width: 18 }, { width: 20 }, { width: 34 }, { width: 22 }, { width: 20 },
    ];

    // Logo, anchored top-left.
    try {
      const imgId = wb.addImage({ base64: LOGO, extension: "png" });
      ws.addImage(imgId, { tl: { col: 0.15, row: 0.2 }, ext: { width: 150, height: 110 } });
    } catch (e) { /* logo optional */ }
    ws.getRow(1).height = 28;
    ws.getRow(2).height = 28;
    ws.getRow(3).height = 28;
    ws.getRow(4).height = 12;

    // Title block (rows 1-3), to the right of the logo.
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

    // Meta block.
    let r = 6;
    const meta = [
      ["Medewerker", name || "—"],
      ["Geldige periode", windowStart && contractEnd ? ${fmtDate(windowStart)} t/m ${fmtDate(contractEnd)} : "—"],
      ["Peildatum (vandaag)", fmtDate(TODAY)],
    ];
    meta.forEach(([k, v]) => {
      const kc = ws.getCell(A${r});
      kc.value = k;
      kc.font = { name: "Arial", size: 10, bold: true, color: { argb: GREY_HEX } };
      const vc = ws.getCell(B${r});
      vc.value = v;
      vc.font = { name: "Arial", size: 11, color: { argb: "FF1A2233" } };
      ws.mergeCells(B${r}:D${r});
      r++;
    });
    r++;

    // Table header.
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

    // Data rows.
    const euroFmt = '€ #,##0.00';
    const dataRows = rows.filter((row) => row.startD && row.endD && row.endD >= row.startD);
    dataRows.forEach((row, idx) => {
      const vals = [
        fmtDate(row.startD),
        fmtDate(row.endD),
        row.monthlyNum,
        row.clamped ? ${fmtDate(row.clamped.start)} t/m ${fmtDate(row.clamped.end)} : "Buiten venster",
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

    // Totals block.
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
      ws.mergeCells(A${r}:D${r});
      const lc = ws.getCell(A${r});
      lc.value = label;
      lc.alignment = { horizontal: "right", vertical: "middle" };
      lc.font = { name: "Arial", size: 10, bold: highlight, color: { argb: highlight ? BLUE_DEEP_HEX : "FF1A2233" } };
      const ec = ws.getCell(E${r});
      ec.value = val;
      ec.numFmt = euroFmt;
      ec.alignment = { horizontal: "right", vertical: "middle" };
      ec.font = { name: "Arial", size: highlight ? 12 : 10, bold: highlight, color: { argb: highlight ? BLUE_DEEP_HEX : "FF1A2233" } };
      if (highlight) {
        ["A", "B", "C", "D", "E"].forEach((col) => {
          ws.getCell(${col}${r}).fill = { type: "pattern", pattern: "solid", fgColor: { argb: LIGHT_HEX } };
        });
      }
      ws.getRow(r).height = highlight ? 22 : 18;
      r++;
    });

    r += 1;
    ws.mergeCells(A${r}:F${r});
    const note = ws.getCell(A${r});
    note.value =
      "Het bedrag 'vooraf op te nemen' betreft toekomstig contractloon en dient bij eerder vertrek te worden terugbetaald.";
    note.font = { name: "Arial", size: 9, italic: true, color: { argb: GREY_HEX } };
    note.alignment = { wrapText: true, vertical: "top" };
    ws.getRow(r).height = 28;

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const safe = (name || "medewerker").replace(/[^\w\-]+/g, "_");
    a.download = studiebudget_${safe}.xlsx;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const hasResult = rows.some((r) => r.budget > 0);

  return (
    <div style={wrap}>
      <style>{@keyframes spin { to { transform: rotate(360deg); } } .spin { animation: spin 0.8s linear infinite; }}</style>
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
            {name ? Resultaat voor ${name} : "Resultaat"}
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
              <div
                style={{
                  ...statValue,
                  fontSize: 34,
                  color: overspent ? "#ffd9d2" : "#fff",
                }}
              >
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
                <div style={splitHint}>
                  Terug te betalen bij eerder vertrek
                </div>
              </div>
            </div>
          )}

          {overspent && (
            <div style={overspentNote}>
              <AlertTriangle size={14} /> Deze medewerker heeft{" "}
              {fmtEUR(used - totalBudget)} méér gebruikt dan het beschikbare
              budget.
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
            style={{
              ...exportBtn,
              opacity: hasResult ? 1 : 0.45,
              cursor: hasResult ? "pointer" : "not-allowed",
            }}
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
const GOLD_DEEP = "#6f6a24";

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
  border: 1px dashed ${BLUE},
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
  border: 1px dashed ${BLUE},
  color: BLUE,
  borderRadius: 9,
  padding: "9px 14px",
  fontSize: 14,
  cursor: "pointer",
  fontWeight: 600,
};
const resultCard = {
  background: linear-gradient(135deg, ${BLUE}, #14306f),
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

Let me check the current state of the repository to understand what needs to be done.

Ran
Check repo state
It looks like the repository only has a README. What would you like me to do with this code? For example:

Add it to the repo — set up a React project and add this component as the main app
Review it — check for bugs, improvements, or security issues
Modify it — make specific changes (tell me what you'd like changed)
What's your goal?

built an app using this code

I'll set up a React app with this component. Let me scaffold the project structure.


Ran 4 commands
Now write the component and wire up the app:


Ran 2 commands
Now write the component file:

Ran
Test partial write
Let me write the full component using Python to avoid shell heredoc limits:


Ran a command, created a file, read a file
Now verify the build works, then commit and push:

Ran
Build the Vite app
Build passes. Now commit and push:


Ran 2 commands
Done. The app is committed and pushed to claude/clever-lovelace-dfkcuw. To run it locally:

cd app
npm install
npm run dev
The build succeeds cleanly. The one warning (chunk > 500 kB) is expected — ExcelJS is large but not a problem for an internal HR tool.

What do I fill in at Main File Path if I want to peploy the app via streamlit.io



