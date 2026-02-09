// Tailwind Configuration
tailwind.config = {
  theme: {
    extend: {
      colors: {
        ecuaplus: {
          orange: "#FF8c00", // Adjusted specific orange
          blue: "#0056b3", // Adjusted specific blue
          darkblue: "#003366",
        },
      },
    },
  },
};

// Application Logic
document.addEventListener("DOMContentLoaded", () => {
  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("file-input");
  const statusArea = document.getElementById("status-area");
  const errorArea = document.getElementById("error-area");

  // Drag & Drop Events
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    dropZone.addEventListener(eventName, highlight, false);
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, unhighlight, false);
  });

  function highlight(e) {
    dropZone.classList.add("drag-active");
  }

  function unhighlight(e) {
    dropZone.classList.remove("drag-active");
  }

  dropZone.addEventListener("drop", handleDrop, false);
  fileInput.addEventListener("change", handleFileSelect, false);

  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleFiles(files);
  }

  function handleFileSelect(e) {
    const files = e.target.files;
    handleFiles(files);
  }

  function handleFiles(files) {
    if (files.length > 0) {
      const file = files[0];
      if (
        file.type === "text/plain" ||
        file.name.toLowerCase().endsWith(".txt")
      ) {
        processFile(file);
      } else {
        showError("Por favor sube un archivo de texto (.txt)");
      }
    }
  }

  function showError(msg) {
    statusArea.classList.add("hidden");
    errorArea.innerHTML = `<div class="p-4 rounded-lg bg-red-50 text-red-700 text-center">${msg}</div>`;
    errorArea.classList.remove("hidden");
  }

  function processFile(file) {
    // Reset UI
    errorArea.classList.add("hidden");
    statusArea.classList.remove("hidden");

    const reader = new FileReader();

    // CRITICAL: ISO-8859-1 encoding
    reader.readAsText(file, "ISO-8859-1");

    reader.onload = function (e) {
      const csvData = e.target.result;

      Papa.parse(csvData, {
        header: true,
        skipEmptyLines: true,
        delimiter: "\t",
        complete: function (results) {
          if (results.errors.length > 0 && results.data.length === 0) {
            showError("Error al leer el archivo. Verifica el formato.");
            return;
          }
          try {
            generateExcel(results.data, file.name);
          } catch (err) {
            showError("Error procesando datos: " + err.message);
            console.error(err);
          }
        },
      });
    };

    reader.onerror = function () {
      showError("Error al leer el archivo.");
    };
  }

  function generateExcel(data, originalFileName) {
    // Critical Fields to keep as String
    const stringFields = [
      "RUC_EMISOR",
      "CLAVE_ACCESO",
      "IDENTIFICACION_RECEPTOR",
    ];
    // Number Fields for calculation
    const numFields = ["VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"];

    // Process Data
    let processedData = data.map((row) => {
      let newRow = { ...row };

      // Enforce Strings
      stringFields.forEach((field) => {
        if (newRow[field]) {
          newRow[field] = String(newRow[field]);
        }
      });

      // Enforce Numbers and Handle Formats (commas vs dots)
      numFields.forEach((field) => {
        if (newRow[field]) {
          let val = String(newRow[field]).trim();
          val = val.replace(",", ".");
          newRow[field] = parseFloat(val) || 0;
        } else {
          newRow[field] = 0;
        }
      });

      return newRow;
    });

    // Calculate Totals
    let totalImporte = 0;
    let totalBase = 0;
    let totalIva = 0;

    let subtotalIvaGrabado = 0;
    let subtotalIvaCero = 0;

    processedData.forEach((row) => {
      const base = row["VALOR_SIN_IMPUESTOS"] || 0;
      const iva = row["IVA"] || 0;
      const total = row["IMPORTE_TOTAL"] || 0;

      totalBase += base;
      totalIva += iva;
      totalImporte += total;

      if (iva > 0) {
        subtotalIvaGrabado += base;
      } else {
        subtotalIvaCero += base;
      }
    });

    // Create Worksheet
    const ws = XLSX.utils.json_to_sheet(processedData);

    // Set column width for better readability
    const wscols = Object.keys(processedData[0] || {}).map((k) => ({
      wch: 15,
    }));
    ws["!cols"] = wscols;

    // Mark specific columns as Text to strictly prevent scientific notation
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const headers = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
      if (cell) headers[C] = cell.v;
    }

    // Iterate over all data cells for critical columns and set type to 's' (string)
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const header = headers[C];
        if (stringFields.includes(header)) {
          const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[cellRef]) continue;
          ws[cellRef].t = "s"; // Force string type
          ws[cellRef].z = "@"; // Force text format
        }
      }
    }

    // Append Totals to the Sheet
    const lastRow = range.e.r + 2;

    // Row 1: Totals of specific columns
    let colBase = headers.indexOf("VALOR_SIN_IMPUESTOS");
    let colIva = headers.indexOf("IVA");
    let colTotal = headers.indexOf("IMPORTE_TOTAL");

    let rowTotals = new Array(headers.length).fill("");
    rowTotals[0] = "TOTALES GENERALES:";
    if (colBase !== -1) rowTotals[colBase] = totalBase;
    if (colIva !== -1) rowTotals[colIva] = totalIva;
    if (colTotal !== -1) rowTotals[colTotal] = totalImporte;

    // Row 2: Subtotal IVA Grabado
    let rowSubGrabado = new Array(headers.length).fill("");
    rowSubGrabado[0] = "Subtotal IVA Grabado (>0%):";
    if (colBase !== -1) rowSubGrabado[colBase] = subtotalIvaGrabado;

    // Row 3: Subtotal IVA 0%
    let rowSubCero = new Array(headers.length).fill("");
    rowSubCero[0] = "Subtotal IVA 0%:";
    if (colBase !== -1) rowSubCero[colBase] = subtotalIvaCero;

    // Append these rows
    XLSX.utils.sheet_add_aoa(
      ws,
      [
        [], // Empty row for spacing
        rowTotals,
        rowSubGrabado,
        rowSubCero,
      ],
      { origin: -1 },
    );

    // Export
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Comprobantes");

    // Generate filename
    const newFileName =
      originalFileName.replace(".txt", "").replace(".TXT", "") +
      "_procesado.xlsx";

    XLSX.writeFile(wb, newFileName);

    // Reset UI
    statusArea.classList.add("hidden");
    errorArea.innerHTML = `<div class="p-4 rounded-lg bg-green-50 text-green-700 text-center font-semibold">¡Archivo procesado y descargado exitosamente!</div>`;
    errorArea.classList.remove("hidden");
  }
});
