// api/salas.js - Backend optimizado y a prueba de errores
export default async function handler(req, res) {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    try {
        // Usamos una URL estándar sin filtros complejos para asegurar la conexión
        const sharepointUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const response = await fetch(sharepointUrl, {
            headers: {
                'Authorization': `Bearer ${process.env.MICROSOFT_GRAPH_TOKEN}`,
                'Accept': 'application/json'
            }
        });

        if (!response.ok) {
            const errorData = await response.text();
            throw new Error(`Error ${response.status}: ${errorData}`);
        }

        const data = await response.json();
        return res.status(200).json(data);

    } catch (error) {
        console.error("Error crítico en API:", error);
        return res.status(500).json({ error: "Fallo de conexión", detalles: error.message });
    }
}
