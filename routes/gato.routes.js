import { Router } from "express";
import { db } from "../config/firebase.js"; // Asegúrate de ajustar esta ruta según dónde importas la instancia de Firestore

const router = Router();

// 1. OBTENER TODOS LOS GATOS (GET /gato)
router.get("/gato", async (req, res) => {
  try {
    const snapshot = await db.collection("gatos").get();

    const gatos = snapshot.docs.map((doc) => ({
      id: doc.id,
      ...doc.data(),
    }));

    return res.status(200).json(gatos);
  } catch (error) {
    console.error("Error al obtener los gatos:", error);
    return res.status(500).json({ mensaje: "Error al obtener la lista de gatos" });
  }
});

// 2. REGISTRAR UN GATO (POST /gato)
// Aquí mantienes la ruta que ya tenías para crear el gato
router.post("/gato", async (req, res) => {
  try {
    // Tu lógica de multer / Supabase / Firestore para registrar
  } catch (error) {
    console.error("Error al registrar el gato:", error);
    return res.status(500).json({ mensaje: "Error al registrar el gato" });
  }
});

export default router;