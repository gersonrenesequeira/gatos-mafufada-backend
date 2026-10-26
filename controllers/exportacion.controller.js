import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";
import { calcularEstadisticas } from "./estadisticas.controller.js";
import db from "../firebase.js";

// ======================================================
// EXPORTAR ESTADÍSTICAS A EXCEL
// ======================================================

export const exportarEstadisticasExcel = async (req, res) => {
  try {
    const stats = await calcularEstadisticas();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Sistema de Gatos";
    workbook.created = new Date();

    // ======================
    // HOJA 1: Resumen
    // ======================
    const hojaResumen = workbook.addWorksheet("Resumen", {
      properties: { tabColor: { argb: "FF4472C4" } },
    });

    const estiloTitulo = {
      font: { bold: true, size: 16, color: { argb: "FFFFFFFF" } },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF4472C4" },
      },
      alignment: { horizontal: "center", vertical: "middle" },
    };

    const estiloEncabezado = {
      font: { bold: true, color: { argb: "FFFFFFFF" } },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF5B9BD5" },
      },
      alignment: { horizontal: "center" },
    };

    const estiloLabel = {
      font: { bold: true },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD6EAF8" },
      },
    };

    hojaResumen.mergeCells("A1:B1");
    hojaResumen.getCell("A1").value = "📊 ESTADÍSTICAS DE GATOS";
    hojaResumen.getCell("A1").style = estiloTitulo;
    hojaResumen.getRow(1).height = 30;

    const datosResumen = [
      ["Total de gatos", stats.total],
      ["Promedio de edad (años)", stats.promedioEdad],
      ["Promedio de peso (kg)", stats.promedioPeso],
      ["Edad mínima", stats.edadMin ?? "—"],
      ["Edad máxima", stats.edadMax ?? "—"],
      ["Peso mínimo (kg)", stats.pesoMin ?? "—"],
      ["Peso máximo (kg)", stats.pesoMax ?? "—"],
      ["Fecha de generación", new Date().toLocaleString("es-ES")],
    ];

    hojaResumen.addRow([]);
    hojaResumen.addRow(["Métrica", "Valor"]).eachCell((cell) => {
      cell.style = estiloEncabezado;
    });

    datosResumen.forEach(([label, valor]) => {
      const row = hojaResumen.addRow([label, valor]);
      row.getCell(1).style = estiloLabel;
    });

    hojaResumen.getColumn(1).width = 28;
    hojaResumen.getColumn(2).width = 20;

    // ======================
    // HOJA 2: Por Raza
    // ======================
    const hojaRazas = workbook.addWorksheet("Por Raza", {
      properties: { tabColor: { argb: "FF70AD47" } },
    });

    hojaRazas.mergeCells("A1:C1");
    hojaRazas.getCell("A1").value = "Distribución por Raza";
    hojaRazas.getCell("A1").style = {
      ...estiloTitulo,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF70AD47" },
      },
    };
    hojaRazas.getRow(1).height = 28;

    hojaRazas.addRow([]);
    hojaRazas
      .addRow(["Raza", "Cantidad", "Porcentaje (%)"])
      .eachCell((cell) => {
        cell.style = {
          ...estiloEncabezado,
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF70AD47" },
          },
        };
      });

    Object.entries(stats.porRaza)
      .sort((a, b) => b[1] - a[1])
      .forEach(([raza, cantidad]) => {
        const porcentaje =
          stats.total > 0
            ? Number(((cantidad / stats.total) * 100).toFixed(1))
            : 0;

        hojaRazas.addRow([raza, cantidad, porcentaje]);
      });

    hojaRazas.getColumn(1).width = 25;
    hojaRazas.getColumn(2).width = 12;
    hojaRazas.getColumn(3).width = 16;

    // ======================
    // HOJA 3: Por Mes
    // ======================
    const hojaMeses = workbook.addWorksheet("Por Mes", {
      properties: { tabColor: { argb: "FFED7D31" } },
    });

    hojaMeses.mergeCells("A1:B1");
    hojaMeses.getCell("A1").value = "Gatos registrados por mes";
    hojaMeses.getCell("A1").style = {
      ...estiloTitulo,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFED7D31" },
      },
    };
    hojaMeses.getRow(1).height = 28;

    hojaMeses.addRow([]);
    hojaMeses.addRow(["Mes (YYYY-MM)", "Cantidad"]).eachCell((cell) => {
      cell.style = {
        ...estiloEncabezado,
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFED7D31" },
        },
      };
    });

    Object.entries(stats.porMes)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([mes, cantidad]) => {
        hojaMeses.addRow([mes, cantidad]);
      });

    hojaMeses.getColumn(1).width = 18;
    hojaMeses.getColumn(2).width = 12;

    // ======================
    // HOJA 4: Listado de Gatos
    // ======================
    const hojaGatos = workbook.addWorksheet("Listado de Gatos", {
      properties: { tabColor: { argb: "FF7030A0" } },
    });

    hojaGatos.mergeCells("A1:F1");
    hojaGatos.getCell("A1").value = "Listado completo de gatos";
    hojaGatos.getCell("A1").style = {
      ...estiloTitulo,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF7030A0" },
      },
    };
    hojaGatos.getRow(1).height = 28;

    hojaGatos.addRow([]);
    hojaGatos
      .addRow(["ID", "Nombre", "Edad", "Peso (kg)", "Raza", "Fecha de registro"])
      .eachCell((cell) => {
        cell.style = {
          ...estiloEncabezado,
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF7030A0" },
          },
        };
      });

    stats.gatos.forEach((g) => {
      hojaGatos.addRow([
        g.id,
        g.nombre,
        g.edad,
        g.peso,
        g.raza,
        g.fecha ? new Date(g.fecha).toLocaleString("es-ES") : "—",
      ]);
    });

    hojaGatos.getColumn(1).width = 28;
    hojaGatos.getColumn(2).width = 18;
    hojaGatos.getColumn(3).width = 10;
    hojaGatos.getColumn(4).width = 12;
    hojaGatos.getColumn(5).width = 18;
    hojaGatos.getColumn(6).width = 22;

    const nombreArchivo = `estadisticas-gatos-${new Date()
      .toISOString()
      .slice(0, 10)}.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nombreArchivo}"`
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Error al generar Excel:", error);
    res.status(500).json({
      mensaje: "Error al generar el archivo Excel.",
      error: error.message,
    });
  }
};

// ======================================================
// EXPORTAR ESTADÍSTICAS A PDF
// ======================================================

export const exportarEstadisticasPDF = async (req, res) => {
  try {
    const stats = await calcularEstadisticas();

    const doc = new PDFDocument({
      size: "A4",
      margin: 50,
      bufferPages: true,
      info: {
        Title: "Estadísticas de Gatos",
        Author: "Sistema de Gatos",
      },
    });

    const nombreArchivo = `estadisticas-gatos-${new Date()
      .toISOString()
      .slice(0, 10)}.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nombreArchivo}"`
    );

    doc.pipe(res);

    const colorPrincipal = "#4472C4";
    const colorSecundario = "#5B9BD5";
    const colorTexto = "#333333";

    // ========= TÍTULO =========

    doc
      .fontSize(22)
      .fillColor(colorPrincipal)
      .font("Helvetica-Bold")
      .text("Estadísticas de Gatos", { align: "center" });

    doc.moveDown(0.3);

    doc
      .fontSize(10)
      .fillColor("#666666")
      .font("Helvetica")
      .text(`Generado el: ${new Date().toLocaleString("es-ES")}`, {
        align: "center",
      });

    doc.moveDown(1.5);

    // ========= RESUMEN =========

    doc
      .fontSize(14)
      .fillColor(colorPrincipal)
      .font("Helvetica-Bold")
      .text("Resumen General");

    doc
      .moveTo(50, doc.y + 5)
      .lineTo(545, doc.y + 5)
      .strokeColor(colorPrincipal)
      .lineWidth(1.5)
      .stroke();

    doc.moveDown(0.8);

    const resumen = [
      ["Total de gatos", stats.total],
      ["Promedio de edad", `${stats.promedioEdad} años`],
      ["Promedio de peso", `${stats.promedioPeso} kg`],
      ["Edad mínima", stats.edadMin !== null ? `${stats.edadMin} años` : "—"],
      ["Edad máxima", stats.edadMax !== null ? `${stats.edadMax} años` : "—"],
      ["Peso mínimo", stats.pesoMin !== null ? `${stats.pesoMin} kg` : "—"],
      ["Peso máximo", stats.pesoMax !== null ? `${stats.pesoMax} kg` : "—"],
    ];

    resumen.forEach(([label, valor]) => {
      doc
        .fontSize(11)
        .fillColor(colorTexto)
        .font("Helvetica-Bold")
        .text(`${label}: `, { continued: true })
        .font("Helvetica")
        .text(String(valor));

      doc.moveDown(0.3);
    });

    doc.moveDown(1);

    // ========= RAZAS =========

    doc
      .fontSize(14)
      .fillColor(colorPrincipal)
      .font("Helvetica-Bold")
      .text("Distribución por Raza");

    doc
      .moveTo(50, doc.y + 5)
      .lineTo(545, doc.y + 5)
      .strokeColor(colorPrincipal)
      .lineWidth(1.5)
      .stroke();

    doc.moveDown(0.8);

    if (Object.keys(stats.porRaza).length === 0) {
      doc.fontSize(11).fillColor("#666").text("No hay datos de razas.");
    } else {
      const startY = doc.y;

      doc
        .rect(50, startY, 495, 22)
        .fill(colorSecundario);

      doc
        .fillColor("#FFFFFF")
        .font("Helvetica-Bold")
        .fontSize(10)
        .text("Raza", 60, startY + 6)
        .text("Cantidad", 300, startY + 6)
        .text("Porcentaje", 420, startY + 6);

      doc.y = startY + 28;

      const razasOrdenadas = Object.entries(stats.porRaza).sort(
        (a, b) => b[1] - a[1]
      );

      razasOrdenadas.forEach(([raza, cantidad], index) => {
        const porcentaje =
          stats.total > 0
            ? ((cantidad / stats.total) * 100).toFixed(1)
            : "0.0";

        const filaY = doc.y;

        if (index % 2 === 0) {
          doc.rect(50, filaY - 4, 495, 20).fill("#F2F2F2");
        }

        doc
          .fillColor(colorTexto)
          .font("Helvetica")
          .fontSize(10)
          .text(raza, 60, filaY)
          .text(String(cantidad), 300, filaY)
          .text(`${porcentaje}%`, 420, filaY);

        doc.y = filaY + 18;
      });
    }

    doc.moveDown(1.2);

    // ========= POR MES =========

    doc
      .fontSize(14)
      .fillColor(colorPrincipal)
      .font("Helvetica-Bold")
      .text("Gatos registrados por mes");

    doc
      .moveTo(50, doc.y + 5)
      .lineTo(545, doc.y + 5)
      .strokeColor(colorPrincipal)
      .lineWidth(1.5)
      .stroke();

    doc.moveDown(0.8);

    if (Object.keys(stats.porMes).length === 0) {
      doc.fontSize(11).fillColor("#666").text("No hay datos por mes.");
    } else {
      Object.entries(stats.porMes)
        .sort((a, b) => a[0].localeCompare(b[0]))
        .forEach(([mes, cantidad]) => {
          doc
            .fontSize(11)
            .fillColor(colorTexto)
            .font("Helvetica-Bold")
            .text(`${mes}: `, { continued: true })
            .font("Helvetica")
            .text(`${cantidad} gato(s)`);

          doc.moveDown(0.3);
        });
    }

    // ========= PIE DE PÁGINA =========

    const rangoPaginas = doc.bufferedPageRange();

    for (let i = 0; i < rangoPaginas.count; i++) {
      doc.switchToPage(i);

      doc
        .fontSize(8)
        .fillColor("#999999")
        .text(
          `Página ${i + 1} de ${rangoPaginas.count} | Sistema de Gatos`,
          50,
          doc.page.height - 40,
          {
            align: "center",
            width: 495,
          }
        );
    }

    doc.end();
  } catch (error) {
    console.error("Error al generar PDF:", error);

    if (!res.headersSent) {
      res.status(500).json({
        mensaje: "Error al generar el archivo PDF.",
        error: error.message,
      });
    }
  }
};

export const exportarGatosLista = async (req, res) => {
  try {
    const snapshot = await db.collection("gatos").orderBy("fecha", "desc").get();

    const gatos = snapshot.docs.map((doc) => ({
      id: doc.id,
      ...doc.data(),
    }));

    if (gatos.length === 0) {
      return res.status(200).type("text/plain").send("No hay gatos registrados.");
    }

    const lineas = gatos.map((g, i) => {
      const fecha = g.fecha
        ? new Date(g.fecha).toLocaleString("es-ES")
        : "—";

      return [
        `${i + 1}. ${g.nombre || "Sin nombre"}`,
        `   ID: ${g.id}`,
        `   Edad: ${g.edad ?? "—"} años`,
        `   Peso: ${g.peso ?? "—"} kg`,
        `   Raza: ${g.raza || "—"}`,
        `   Fecha: ${fecha}`,
        `   Imagen: ${g.imagenUrl || "—"}`,
        "",
      ].join("\n");
    });

    const contenido =
      `LISTADO DE GATOS (${gatos.length})\n` +
      `Generado: ${new Date().toLocaleString("es-ES")}\n` +
      `${"─".repeat(40)}\n\n` +
      lineas.join("\n");

    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.setHeader("Content-Disposition", "inline; filename=\"gatos-lista.txt\"");
    res.send(contenido);
  } catch (error) {
    console.error("Error al exportar lista:", error);
    res.status(500).json({
      mensaje: "Error al generar el listado de gatos.",
      error: error.message,
    });
  }
};

export const exportarGatosJSON = async (req, res) => {
  try {
    const snapshot = await db.collection("gatos").orderBy("fecha", "desc").get();

    const gatos = snapshot.docs.map((doc) => ({
      id: doc.id,
      ...doc.data(),
    }));

    const jsonFormateado = JSON.stringify(gatos, null, 2);

    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.setHeader("Content-Disposition", "inline; filename=\"gatos.json\"");
    res.send(jsonFormateado);
  } catch (error) {
    console.error("Error al exportar JSON:", error);
    res.status(500).json({
      mensaje: "Error al generar el JSON de gatos.",
      error: error.message,
    });
  }
};

export const exportarGatosSQL = async (req, res) => {
  try {
    const snapshot = await db.collection("gatos").orderBy("fecha", "desc").get();

    const gatos = snapshot.docs.map((doc) => ({
      id: doc.id,
      ...doc.data(),
    }));

    const escaparSQL = (valor) => {
      if (valor === null || valor === undefined) return "NULL";
      if (typeof valor === "number") return valor;
      return `'${String(valor).replace(/'/g, "''")}'`;
    };

    let sql = `-- Script de inserción de gatos\n`;
    sql += `-- Generado: ${new Date().toLocaleString("es-ES")}\n`;
    sql += `-- Total registros: ${gatos.length}\n\n`;

    sql += `CREATE TABLE IF NOT EXISTS gatos (\n`;
    sql += `  id VARCHAR(50) PRIMARY KEY,\n`;
    sql += `  nombre VARCHAR(100) NOT NULL,\n`;
    sql += `  edad INT,\n`;
    sql += `  peso DECIMAL(5,2),\n`;
    sql += `  raza VARCHAR(80),\n`;
    sql += `  imagen_url TEXT,\n`;
    sql += `  fecha DATETIME\n`;
    sql += `);\n\n`;

    if (gatos.length === 0) {
      sql += `-- No hay registros para insertar\n`;
    } else {
      gatos.forEach((g) => {
        const fecha = g.fecha
          ? new Date(g.fecha).toISOString().slice(0, 19).replace("T", " ")
          : null;

        sql += `INSERT INTO gatos (id, nombre, edad, peso, raza, imagen_url, fecha) VALUES (\n`;
        sql += `  ${escaparSQL(g.id)},\n`;
        sql += `  ${escaparSQL(g.nombre)},\n`;
        sql += `  ${escaparSQL(g.edad)},\n`;
        sql += `  ${escaparSQL(g.peso)},\n`;
        sql += `  ${escaparSQL(g.raza)},\n`;
        sql += `  ${escaparSQL(g.imagenUrl)},\n`;
        sql += `  ${escaparSQL(fecha)}\n`;
        sql += `);\n\n`;
      });
    }

    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.setHeader("Content-Disposition", "inline; filename=\"gatos-insert.sql\"");
    res.send(sql);
  } catch (error) {
    console.error("Error al exportar SQL:", error);
    res.status(500).json({
      mensaje: "Error al generar el script SQL.",
      error: error.message,
    });
  }
};