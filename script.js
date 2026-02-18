// Configuración de Tailwind
tailwind.config = {
  theme: {
    extend: {
      colors: {
        ecuaplus: {
          orange: "#FF8c00", // Naranja específico ajustado
          blue: "#0056b3", // Azul específico ajustado
          darkblue: "#003366",
        },
      },
    },
  },
};

// Lógica de la Aplicación
document.addEventListener("DOMContentLoaded", () => {
  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("file-input");
  const statusArea = document.getElementById("status-area");
  const errorArea = document.getElementById("error-area");

  // Eventos de Arrastrar y Soltar (Drag & Drop)
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
    const validFiles = [];
    const invalidFiles = [];

    // Iterar sobre la lista de archivos
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (file.type === "text/plain" || file.name.toLowerCase().endsWith(".txt")) {
        validFiles.push(file);
      } else {
        invalidFiles.push(file.name);
      }
    }

    if (invalidFiles.length > 0) {
      showError(`Archivos ignorados (no son .txt): ${invalidFiles.join(", ")}`);
      // Aún procesamos los archivos válidos si existen
    }

    if (validFiles.length > 0) {
      processFiles(validFiles);
    } else if (invalidFiles.length === 0 && files.length > 0) {
      showError("Por favor sube archivos de texto (.txt)");
    }
  }

  function showError(msg) {
    statusArea.classList.add("hidden");
    errorArea.innerHTML = `<div class="p-4 rounded-lg bg-red-50 text-red-700 text-center">${msg}</div>`;
    errorArea.classList.remove("hidden");
  }

  // Envolver FileReader y PapaParse en una Promesa
  function readAndParseFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsText(file, "ISO-8859-1");

      reader.onload = function (e) {
        const csvData = e.target.result;
        Papa.parse(csvData, {
          header: true,
          skipEmptyLines: true,
          delimiter: "\t",
          complete: function (results) {
            if (results.errors.length > 0 && results.data.length === 0) {
              console.warn(`Error parseando ${file.name}`, results.errors);
              resolve({ data: [], filename: file.name, error: "Formato inválido o archivo vacío" });
            } else {
              resolve({ data: results.data, filename: file.name });
            }
          },
          error: function (err) {
            reject(err);
          }
        });
      };

      reader.onerror = function () {
        reject(new Error(`Error leyendo archivo: ${file.name}`));
      };
    });
  }


  async function processFiles(files) {
    // Reiniciar interfaz
    errorArea.classList.add("hidden");
    statusArea.classList.remove("hidden");
    statusArea.innerHTML = `
        <div class="flex items-center justify-center p-4 rounded-lg bg-blue-50 text-blue-800">
            <svg class="animate-spin -ml-1 mr-3 h-5 w-5 text-blue-800" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
            </svg>
            Procesando ${files.length} archivo(s)...
        </div>`;


    try {
      const promises = files.map(file => readAndParseFile(file));
      const results = await Promise.all(promises);

      let processedCount = 0;
      let errors = [];

      for (const result of results) {
        if (result.error) {
          errors.push(`${result.filename}: ${result.error}`);
          continue;
        }

        if (result.data.length === 0) {
          errors.push(`${result.filename}: Sin datos`);
          continue;
        }

        try {
          generateExcel(result.data, result.filename);
          processedCount++;
        } catch (err) {
          errors.push(`${result.filename}: Error generando Excel (${err.message})`);
        }
      }

      if (processedCount === 0 && errors.length > 0) {
        showError(`No se pudo procesar ningún archivo. Errores: <br>${errors.join("<br>")}`);
      } else if (errors.length > 0) {
        statusArea.classList.add("hidden");
        errorArea.innerHTML = `
                <div class="p-4 rounded-lg bg-yellow-50 text-yellow-800 text-center">
                    ¡${processedCount} archivo(s) procesado(s) y descargado(s) exitosamente!<br>
                    <span class="text-sm">Advertencia: Algunos archivos tuvieron errores:<br>${errors.join("<br>")}</span>
                </div>`;
        errorArea.classList.remove("hidden");
      } else {
        statusArea.classList.add("hidden");
        errorArea.innerHTML = `<div class="p-4 rounded-lg bg-green-50 text-green-700 text-center font-semibold">¡${processedCount} archivo(s) procesado(s) y descargado(s) exitosamente!</div>`;
        errorArea.classList.remove("hidden");
      }


    } catch (err) {
      showError("Error procesando archivos: " + err.message);
      console.error(err);
    }
  }

  function generateExcel(data, originalFileName) {
    // Campos críticos para mantener como Texto
    const stringFields = [
      "RUC_EMISOR",
      "CLAVE_ACCESO",
      "IDENTIFICACION_RECEPTOR",
    ];
    // Campos numéricos
    const numFields = ["VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"];

    // Procesar Datos
    let processedData = data.map((row) => {
      let newRow = { ...row };

      stringFields.forEach((field) => {
        if (newRow[field]) {
          newRow[field] = String(newRow[field]);
        }
      });

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

    // Calcular Totales
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

    // Crear Hoja de Cálculo
    const ws = XLSX.utils.json_to_sheet(processedData);

    // --- ESTILOS ---

    // Definir objetos de estilo
    const headerStyle = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "003366" } }, // Azul oscuro institucional
      alignment: { horizontal: "center" },
      border: {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } }
      }
    };

    const cellStyle = {
      border: {
        top: { style: "thin", color: { rgb: "CCCCCC" } },
        bottom: { style: "thin", color: { rgb: "CCCCCC" } },
        left: { style: "thin", color: { rgb: "CCCCCC" } },
        right: { style: "thin", color: { rgb: "CCCCCC" } }
      }
    };

    const currencyStyle = {
      ...cellStyle,
      numFmt: '"$"#,##0.00' // Formato moneda
    };

    // Aplicar estilos
    const range = XLSX.utils.decode_range(ws["!ref"]);

    // Ancho de columnas (Auto-width aproximado)
    const wscols = [];

    // 1. Obtener encabezados y aplicar estilo
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
      if (!ws[cellAddress]) continue;

      // Aplicar estilo de encabezado
      ws[cellAddress].s = headerStyle;

      // Calcular ancho basado en la longitud del encabezado + padding
      let colWidth = (ws[cellAddress].v ? String(ws[cellAddress].v).length : 10) + 5;
      wscols.push({ wch: colWidth });
    }
    ws["!cols"] = wscols;

    // 2. Aplicar estilos a las celdas de datos
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellAddress]) continue;

        const headerCell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
        const headerName = headerCell ? headerCell.v : "";

        // Determinar estilo base
        let styleToApply = { ...cellStyle };

        // Si es columna numérica, aplicar formato moneda
        if (numFields.includes(headerName)) {
          styleToApply = { ...currencyStyle };
        }

        // Si es campo de texto crítico, forzar tipo string y formato texto
        if (stringFields.includes(headerName)) {
          ws[cellAddress].t = "s";
          ws[cellAddress].z = "@";
          styleToApply.alignment = { horizontal: "left" };
        }

        ws[cellAddress].s = styleToApply;
      }
    }

    // --- FIN ESTILOS ---

    // Añadir Totales
    const lastRowIndex = range.e.r + 2;

    // Obtener encabezados para identificar columnas
    const headers = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
      headers.push(cell ? cell.v : "");
    }

    // Fila 1: Totales de columnas específicas
    let colBase = headers.indexOf("VALOR_SIN_IMPUESTOS");
    let colIva = headers.indexOf("IVA");
    let colTotal = headers.indexOf("IMPORTE_TOTAL");

    let rowTotals = new Array(headers.length).fill("");
    rowTotals[0] = "TOTALES GENERALES:";
    if (colBase !== -1) rowTotals[colBase] = totalBase;
    if (colIva !== -1) rowTotals[colIva] = totalIva;
    if (colTotal !== -1) rowTotals[colTotal] = totalImporte;

    // Fila 2: Subtotal IVA Grabado
    let rowSubGrabado = new Array(headers.length).fill("");
    rowSubGrabado[0] = "Subtotal IVA Grabado (>0%):";
    if (colBase !== -1) rowSubGrabado[colBase] = subtotalIvaGrabado;

    // Fila 3: Subtotal IVA 0%
    let rowSubCero = new Array(headers.length).fill("");
    rowSubCero[0] = "Subtotal IVA 0%:";
    if (colBase !== -1) rowSubCero[colBase] = subtotalIvaCero;

    // Añadir estas filas al final (automáticamente actualiza el rango !ref)
    XLSX.utils.sheet_add_aoa(
      ws,
      [
        [], // Fila vacía para espaciado
        rowTotals,
        rowSubGrabado,
        rowSubCero,
      ],
      { origin: -1 }
    );

    // --- APLICAR ESTILOS A LOS TOTALES ---
    // Recalcular el rango final después de añadir las filas
    const finalRange = XLSX.utils.decode_range(ws["!ref"]);
    const startRowTotals = lastRowIndex; // Fila donde empiezan los totales (después de la vacía)

    // Estilo para totales
    const totalStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "E6F0FF" } }, // Azul muy claro
      border: {
        top: { style: "thin" }, bottom: { style: "thin" },
        left: { style: "thin" }, right: { style: "thin" }
      }
    };

    const totalCurrencyStyle = {
      ...totalStyle,
      numFmt: '"$"#,##0.00'
    };

    // Iterar sobre las filas de totales añadidas
    for (let R = startRowTotals; R <= finalRange.e.r; ++R) {
      for (let C = finalRange.s.c; C <= finalRange.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellAddress]) continue;

        const cellValue = ws[cellAddress].v;
        let styleToApply = { ...totalStyle };

        // Si es número, aplicar formato moneda
        if (typeof cellValue === 'number') {
          styleToApply = { ...totalCurrencyStyle };
        }

        ws[cellAddress].s = styleToApply;
      }
    }

    // Exportar
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Comprobantes");
    const newFileName = originalFileName.replace(/\.txt$/i, "") + "_procesado.xlsx";

    XLSX.writeFile(wb, newFileName);
  }
});
