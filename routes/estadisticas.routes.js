import express from "express";
import { obtenerEstadisticas } from "../controllers/estadisticas.controller.js";
import { exportarEstadisticasExcel } from "../controllers/exportacion.controller.js";
import { capturarInfoSolicitud } from "../middlewares/auditoria.js";
import { exportarEstadisticasExcel, exportarEstadisticasPDF } from "../controllers/exportacion.controller.js";

//Las demás importaciones y rutas

// Descarga el PDF


const router = express.Router();

router.use(capturarInfoSolicitud);

router.get("/estadisticas/pdf", exportarEstadisticasPDF);

// Estadísticas en JSON
router.get("/estadisticas", obtenerEstadisticas);

// Descarga el archivo Excel
router.get("/estadisticas/excel", exportarEstadisticasExcel);

export default router;
