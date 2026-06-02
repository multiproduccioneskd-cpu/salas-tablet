// api/salas.js
export default async function handler(req, res) {
    try {
        const url = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${process.env.MICROSOFT_GRAPH_TOKEN}`, 'Accept': 'application/json' }
        });

        const data = await response.json();
        
        // Si hay error en la respuesta de Graph, lo mandamos como JSON para que el HTML lo vea
        if (!response.ok) return res.status(500).json({ error: "Graph API Error", details: data });

        return res.status(200).json(data);
    } catch (error) {
        return res.status(500).json({ error: "Backend Error", details: error.message });
    }
}
