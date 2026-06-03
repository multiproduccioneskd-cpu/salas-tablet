// api/salas.js - Backend actualizado
export default async function handler(req, res) {
    // Permitir llamadas desde tu propio dominio frontend
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET');

    const TENANT_ID = "05970e72-c674-4f1f-8033-6e35dd7f76aa";
    const CLIENT_ID = "ceee9a3e-aa63-419c-960a-321e8726fd65";
    const CLIENT_SECRET = "Lfk8Q~-8lvREUP6Amzkd_7mdAT4Z1o16OdF8PazH";
    const SITE_ID = "ugmchile-my.sharepoint.com,0c9c826e-4733-43c6-a116-fc1ac6ce17b6,842618f0-6656-4eff-aac2-9955aec22a9d";
    const LIST_ID = "c546cdce-816e-4b01-9484-1e41902ee91a";

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
        const TOKEN = tokenData.access_token;

        if (!TOKEN) throw new Error("No se pudo generar el token de Graph API: " + JSON.stringify(tokenData));

        // 2. Consultar los elementos de Sharepoint
        // Se añade 'Prefer' para forzar la lectura de elementos no indexados aún
        const graphResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields&$orderby=createdDateTime desc&$top=100`, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${TOKEN}`,
                'Prefer': 'HonorNonIndexedQueries' 
            }
        });
        
        if (!graphResponse.ok) {
            const errorData = await graphResponse.text();
            throw new Error(`Graph API error ${graphResponse.status}: ${errorData}`);
        }

        const graphData = await graphResponse.json();

        // 3. Responder los datos al frontend
        return res.status(200).json(graphData);

    } catch (error) {
        console.error("Error en /api/salas:", error.message);
        return res.status(500).json({ error: error.message });
    }
}
