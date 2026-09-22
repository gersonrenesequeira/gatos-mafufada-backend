import express from "express";
import multer from "multer";

import {
    registrarGato,
    obtenerGatos,
    actualizarGato
} from "../controllers/gato.controller.js";

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });


// Registrar gato
router.post("/gato", upload.single("imagen"), registrarGato);


// Actualizar gato
router.put("/gato/:id", upload.single("imagen"), actualizarGato);


// Obtener todos los gatos
router.get("/gatos", obtenerGatos);


export default router;