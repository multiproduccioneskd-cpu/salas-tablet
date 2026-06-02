export default async function handler(req, res) {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    try {
        const tokenResponse = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                client_id: process.env.CLIENT_ID,
                client_secret: process.env.CLIENT_SECRET,
                scope: 'https://graph.microsoft.com/.default',
                grant_type: 'client_credentials'
            })
        });

        const tokenData = await tokenResponse.json();
        if (!tokenData.access_token) {
            return res.status(500).json({ error: "Fallo al obtener token", detalles: tokenData });
        }

        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/lists/${process.env.SHAREPOINT_LIST_ID}/items?expand=fields`;
        
        const graphResponse = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${tokenData.access_token}` }
        });

        // AQUÍ ESTÁ EL CAMBIO:
        // Si no es un 200, leemos el texto del error en lugar de intentar parsear JSON
        if (!graphResponse.ok) {
            const errorText = await graphResponse.text();
            return res.status(500).json({ error: "Graph API Error", detalles: errorText });
        }

        const graphData = await graphResponse.json();
        return res.status(200).json(graphData);

    } catch (error) {
        return res.status(500).json({ error: "Error de servidor", detalles: error.message });
    }
}
