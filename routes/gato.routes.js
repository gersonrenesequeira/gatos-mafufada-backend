import express from "express";

import {
    registrarGato,
    obtenerGatos
} from "../controllers/gato.controller.js";

const router = express.Router();


// Registrar gato
router.post("/gato", registrarGato);


// Obtener todos los gatos
router.get("/gatos", obtenerGatos);


export default router;