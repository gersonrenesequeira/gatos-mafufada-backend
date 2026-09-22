import express from "express";
import multer from "multer";
import {
    registrarGato,
    obtenerGatos,
    actualizarGato
} from "../controllers/gato.controller.js";

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });

// Ruta de prueba
router.get("/", (req, res) => {
    res.json({ mensaje: "Servidor de Gatos activo y funcionando correctamente." });
});

// Registrar gato
router.post("/gato", upload.single("imagen"), registrarGato);

// Obtener todos los gatos (agregamos ambos endpoints para evitar descalces)
router.get("/gato", obtenerGatos);
router.get("/gatos", obtenerGatos);

// Actualizar gato
router.put("/gato/:id", upload.single("imagen"), actualizarGato);

export default router;