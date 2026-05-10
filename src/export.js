// ─── EXPORT UTILITIES ─────────────────────────────────────────────────────────

export const exportToExcel = (data, filename, sheetName = "Dados") => {
  try {
    const XLSX = window.XLSX;
    if (!XLSX) { alert("Biblioteca Excel não carregada. Tente recarregar a página."); return; }
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    // Auto column width
    const maxW = data.reduce((acc, row) => {
      Object.keys(row).forEach((k, i) => {
        const len = Math.max((acc[i] || 10), String(row[k] || "").length + 2, k.length + 2);
        acc[i] = Math.min(len, 40);
      });
      return acc;
    }, {});
    ws["!cols"] = Object.values(maxW).map(w => ({ wch: w }));
    XLSX.writeFile(wb, `${filename}.xlsx`);
  } catch (e) {
    console.error("Excel export error:", e);
    alert("Erro ao exportar Excel. Verifique o console.");
  }
};

export const exportToPDF = (columns, rows, filename, title) => {
  try {
    // Use dynamic script loading for jsPDF
    const { jsPDF } = window.jspdf;
    if (!jsPDF) { alert("Biblioteca PDF não carregada."); return; }
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    // Header
    doc.setFillColor(15, 17, 26);
    doc.rect(0, 0, 297, 20, "F");
    doc.setTextColor(232, 234, 246);
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.text("CaixaPro · Tirzepatida", 14, 13);
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(180, 180, 200);
    doc.text(title, 100, 13);
    doc.text(new Date().toLocaleString("pt-BR"), 240, 13);
    // Table
    doc.autoTable({
      head: [columns.map(c => c.label)],
      body: rows.map(r => columns.map(c => r[c.key] ?? "")),
      startY: 25,
      styles: { fontSize: 8, cellPadding: 3, textColor: [50, 50, 70] },
      headStyles: { fillColor: [79, 94, 240], textColor: [255, 255, 255], fontStyle: "bold" },
      alternateRowStyles: { fillColor: [245, 245, 252] },
      margin: { left: 14, right: 14 },
    });
    doc.save(`${filename}.pdf`);
  } catch (e) {
    console.error("PDF export error:", e);
    alert("Erro ao exportar PDF.");
  }
};
