import express from "express";
import multer from "multer";

import {
  registrarGato,
  obtenerGatos,
  actualizarGato,
  eliminarGato,
  buscarGatos
} from "../controllers/gato.controller.js";
import { capturarInfoSolicitud } from "../middlewares/auditoria.js";



const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });
router.use(capturarInfoSolicitud);   // aplica a todas las rutas de este router

// Ruta de prueba
router.get("/", (req, res) => {
  res.json({ mensaje: "Servidor de Gatos activo y funcionando correctamente." });
});

// Buscar gatos (debe ir antes de rutas con parametros dinámicos)
router.get("/gatos/buscar", buscarGatos);

// Registrar gato
router.post("/gato", upload.single("imagen"), registrarGato);

// Obtener todos los gatos
router.get("/gato", obtenerGatos);
router.get("/gatos", obtenerGatos);

// Actualizar gato
router.put("/gato/:id", upload.single("imagen"), actualizarGato);

// Eliminar gato
router.delete("/gato/:id", eliminarGato);

export default router;