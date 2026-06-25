// scripts/fetch-onboarding-data.js
// Este script obtiene las carpetas y datos de OneDrive y los convierte en JSON

const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");
const fs = require("fs");

// Configuración
const DRIVE_ID = "your-drive-id"; // Reemplazar con tu Drive ID
const ROOT_FOLDER_PATH = "/Documentos/Talento Humano 4 DESEMGAMOS CON INSPIRACIÓN";

// Crear cliente de Graph
const client = Client.init({
  authProvider: (done) => {
    done(null, process.env.ONEDRIVE_ACCESS_TOKEN);
  },
});

async function fetchFolderStructure(folderId, depth = 0) {
  try {
    const items = await client
      .api(`/drives/${DRIVE_ID}/items/${folderId}/children`)
      .get();

    const structure = [];

    for (const item of items.value) {
      const element = {
        nombre: item.name,
        tipo: item.folder ? "carpeta" : "archivo",
        id: item.id,
        tamaño: item.size || 0,
        fechaModificacion: item.lastModifiedDateTime,
      };

      // Si es carpeta, recursivamente obtener contenido
      if (item.folder && depth < 3) {
        element.contenido = await fetchFolderStructure(item.id, depth + 1);
      }

      structure.push(element);
    }

    return structure;
  } catch (error) {
    console.error("❌ Error al obtener carpetas:", error.message);
    return [];
  }
}

async function main() {
  try {
    console.log("🔄 Obteniendo estructura de OneDrive...");

    // Obtener el ID de la carpeta raíz
    const rootFolder = await client
      .api(`/drives/${DRIVE_ID}/root`)
      .get();

    // Obtener estructura completa
    const estructura = await fetchFolderStructure(rootFolder.id);

    // Crear objeto JSON limpio
    const datosOnboarding = {
      fecha_generacion: new Date().toISOString(),
      ultima_actualizacion: new Date().toLocaleString("es-CO"),
      total_carpetas: contarCarpetas(estructura),
      areas: estructura,
    };

    // Guardar JSON
    fs.writeFileSync(
      "onboarding-raw.json",
      JSON.stringify(datosOnboarding, null, 2)
    );

    console.log("✅ JSON generado correctamente");
    console.log(`📊 Total de carpetas: ${datosOnboarding.total_carpetas}`);

  } catch (error) {
    console.error("❌ Error en la rutina:", error);
    process.exit(1);
  }
}

function contarCarpetas(items) {
  let count = 0;
  for (const item of items) {
    if (item.tipo === "carpeta") {
      count++;
      if (item.contenido) count += contarCarpetas(item.contenido);
    }
  }
  return count;
}

main();
