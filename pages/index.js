"use client";
import { useState } from "react";

const SUBSIDIARY_ORDER = [
  "Ava Labs, Inc.",
  "Antarctica, Inc.",
  "AVL Services LLC",
  "AVL Canada, Inc.",
  "AVL Services (Cayman) SEZC",
  "Karilanche LLC",
  "TMLOC, LLC",
  "Enclave Markets Inc.",
  "Otto Teknoloji Anonim Sirketi (JSC)",
  "Enclave Global, Inc.",
  "Elimination - Ava Labs, Inc.",
  "Elimination - TMLOC",
  "Elimination - Enclave Markets Inc."
];

export default function Home() {
  const [file, setFile] = useState(null);
  const [url, setUrl] = useState(null);

  const handleChange = (e) => {
    const f = e.target.files[0];
    if (f && f.name.endsWith(".xlsx")) setFile(f);
    else alert("Upload a valid .xlsx file");
  };

  const handleProcess = async () => {
    if (!file) return;
    const XLSX = await import("xlsx");

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data);
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const headerIdx = json.findIndex((r) => r.includes("Account"));
      if (headerIdx === -1) throw new Error("Header row not found");

      const headers = json[headerIdx];
      const body = json.slice(headerIdx + 1);

      const subIdx = 0;
      const acctIdx = headers.indexOf("Account");
      const debitIdx = headers.indexOf("Debit");
      const creditIdx = headers.indexOf("Credit");

      const pivot = {};

      for (const row of body) {
        const sub = (row[subIdx] || "").trim();
        const acct = (row[acctIdx] || "").trim();
        if (!acct || !sub) continue;

        const key = acct;
        const val = (parseFloat(row[debitIdx] || 0) - parseFloat(row[creditIdx] || 0));

        if (!pivot[key]) pivot[key] = {};
        if (!pivot[key][sub]) pivot[key][sub] = 0;
        pivot[key][sub] += val;
      }

      const output = [["Account", ...SUBSIDIARY_ORDER, "Total"]];

      for (const acct of Object.keys(pivot)) {
        const row = [acct];
        let total = 0;
        for (const sub of SUBSIDIARY_ORDER) {
          const val = pivot[acct][sub] || 0;
          row.push(val);
          total += val;
        }
        row.push(total);
        output.push(row);
      }

      const ws = XLSX.utils.aoa_to_sheet(output);
      const outWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outWb, ws, "Pivot");

      const buf = XLSX.write(outWb, { type: "array", bookType: "xlsx" });
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const outUrl = URL.createObjectURL(blob);
      setUrl(outUrl);
    } catch (err) {
      alert("Error processing pivot file");
      console.error("Pivot error:", err);
    }
  };

  return (
    <main style={{ padding: 40, fontFamily: "sans-serif" }}>
      <h1 style={{ fontSize: "28px", fontWeight: "bold" }}>NetSuite Pivot Formatter</h1>
      <input type="file" accept=".xlsx" onChange={handleChange} />
      <button onClick={handleProcess} disabled={!file} style={{ marginLeft: 10 }}>
        Process Pivot
      </button>
      {url && (
        <div style={{ marginTop: 20 }}>
          <a href={url} download="pivot_output.xlsx">
            Download Pivot File
          </a>
        </div>
      )}
    </main>
  );
}
