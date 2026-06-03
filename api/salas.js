// api/salas.js - Backend optimizado para lectura amplia
export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET');

    // Usamos variables de entorno para seguridad en GitHub
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

        // 2. Consultar los elementos de Sharepoint
        // Hemos quitado el '$orderby' para evitar que SharePoint filtre por fechas indexadas
        // Hemos dejado '$top=100' y añadido 'HonorNonIndexedQueries' para máxima visibilidad
        const graphResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields&$top=100`, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${tokenData.access_token}`,
                'Prefer': 'HonorNonIndexedQueries' 
            }
        });
        
        if (!graphResponse.ok) {
            const errorText = await graphResponse.text();
            throw new Error(`Graph API error: ${errorText}`);
        }

        const graphData = await graphResponse.json();
        return res.status(200).json(graphData);

    } catch (error) {
        console.error("Error en /api/salas:", error.message);
        return res.status(500).json({ error: error.message });
    }
}
