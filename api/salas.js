// api/salas.js - Backend optimizado para Vercel
export default async function handler(req, res) {
    // Permitir CORS por si acaso
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    const TENANT_ID = process.env.TENANT_ID;
    const CLIENT_ID = process.env.CLIENT_ID;
    const CLIENT_SECRET = process.env.CLIENT_SECRET;
    const SITE_ID = process.env.SHAREPOINT_SITE_ID;
    const LIST_ID = process.env.SHAREPOINT_LIST_ID;

    try {
        // 1. Obtener Token
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
        if (!tokenData.access_token) throw new Error("No se pudo obtener el token");

        // 2. Consultar SharePoint (URL LIMPIA para evitar errores 400)
        // Eliminamos el orderby aquí para asegurar que la API no falle
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields&$top=100`;

        const graphResponse = await fetch(graphUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${tokenData.access_token}` }
        });

        if (!graphResponse.ok) throw new Error("Error al consultar SharePoint");

        const graphData = await graphResponse.json();

        // 3. Responder
        return res.status(200).json(graphData);

    } catch (error) {
        console.error("Error Backend:", error);
        return res.status(500).json({ error: error.message });
    }
}
