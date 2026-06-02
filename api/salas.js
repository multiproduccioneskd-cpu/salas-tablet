// api/salas.js - EL BACKEND ESTÁNDAR QUE SÍ FUNCIONA
export default async function handler(req, res) {
    // Desactivamos el caché para asegurar que traiga lo último de SharePoint
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    try {
        // Esta es la URL limpia que no debería tirar error 400 ni 500
        const sharepointUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const response = await fetch(sharepointUrl, {
            headers: {
                'Authorization': `Bearer ${process.env.MICROSOFT_GRAPH_TOKEN}`,
                'Accept': 'application/json'
            }
        });

        if (!response.ok) {
            console.error("Error desde Microsoft:", await response.text());
            throw new Error(`Estado de respuesta: ${response.status}`);
        }

        const data = await response.json();
        
        // Enviamos la data cruda al frontend
        return res.status(200).json(data);

    } catch (error) {
        console.error("Error en backend:", error);
        return res.status(500).json({ error: "Fallo en la conexión", detalles: error.message });
    }
}
