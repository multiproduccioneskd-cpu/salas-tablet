// api/salas.js - Backend optimizado para lectura de eventos
export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET');

    // IMPORTANTE: En producción/GitHub, estas variables deben venir de process.env
    // Nunca subas secretos reales a GitHub.
    const TENANT_ID = process.env.AZURE_TENANT_ID;
    const CLIENT_ID = process.env.AZURE_CLIENT_ID;
    const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
    const SITE_ID = process.env.SITE_ID; 
    const LIST_ID = process.env.LIST_ID;

    try {
        // 1. Obtener Token de Acceso
        const tokenResponse = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                client_id: CLIENT_ID,
                client_secret: CLIENT_SECRET,
                scope: 'https://graph.microsoft.com/.default',
                grant_type: 'client_credentials'
            })
        });
        
        const tokenData = await tokenResponse.json();
        if (!tokenData.access_token) throw new Error("Error al obtener token");

        // 2. Consultar eventos
        // Usamos el endpoint /events que es el correcto para calendarios
        // Esto captura los cambios en tiempo real mejor que /items
        const graphResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/events?$orderby=start/dateTime desc&$top=50`, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${tokenData.access_token}`,
                'Prefer': 'HonorNonIndexedQueries' // Fuerza lectura sin esperar al índice de búsqueda
            }
        });
        
        if (!graphResponse.ok) {
            const errorText = await graphResponse.text();
            throw new Error(`Graph API error: ${errorText}`);
        }

        const graphData = await graphResponse.json();
        return res.status(200).json(graphData);

    } catch (error) {
        return res.status(500).json({ error: error.message });
    }
}
