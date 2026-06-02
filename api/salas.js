// salas.js - Servidor Backend (Vercel Serverless Function)

export default async function handler(req, res) {
    // Configuramos cabeceras para desactivar el caché duro en Vercel
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');

    try {
        // ID de tu sitio e ID de tu lista de SharePoint de Salas UGM
        const SITE_ID = process.env.SHAREPOINT_SITE_ID;
        const LIST_ID = process.env.SHAREPOINT_LIST_ID;
        const ACCESS_TOKEN = process.env.MICROSOFT_GRAPH_TOKEN; // O tu lógica de autenticación activa

        // QUERY INVERTIDA: Agregamos '$orderby=createdDateTime desc' o 'fields/id desc' 
        // Esto obliga a Microsoft Graph / SharePoint a entregar primero los elementos más nuevos creados en la tarde
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields&$orderby=createdDateTime desc&$top=100`;

        const response = await fetch(graphUrl, {
            headers: {
                'Authorization': `Bearer ${ACCESS_TOKEN}`,
                'Accept': 'application/json',
                'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly' // Permite orderby en campos no indexados si es necesario
            }
        });

        if (!response.ok) {
            throw new Error(`Error en SharePoint API: ${response.statusText}`);
        }

        const data = await response.json();

        // Respondemos directo al HTML con el JSON invertido
        return res.status(200).json(data);

    } catch (error) {
        console.error("Error en backend salas.js:", error);
        return res.status(500).json({ error: "Error interno al conectar con SharePoint", detalles: error.message });
    }
}
