// api/salas.js - Backend seguro y con ajuste de hora
export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET');

    const TENANT_ID = process.env.AZURE_TENANT_ID;
    const CLIENT_ID = process.env.AZURE_CLIENT_ID;
    const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
    // IMPORTANTE: Asegúrate de actualizar SITE_ID y LIST_ID en tu .env/Vercel
    const SITE_ID = process.env.SITE_ID;
    const LIST_ID = process.env.LIST_ID;

    try {
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

        // Consultamos la lista. Quitamos el filtro de fecha para no perder eventos nuevos
        const graphResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields&$top=100`, {
            method: 'GET',
            headers: { 
                'Authorization': `Bearer ${TOKEN}`,
                'Prefer': 'HonorNonIndexedQueries'
            }
        });
        
        const graphData = await graphResponse.json();

        // OPCIONAL: Si necesitas ajustar la hora aquí antes de enviarla al frontend:
        if (graphData.value) {
            graphData.value.forEach(item => {
                if (item.fields.HoraInicio) {
                    // Esto envía la fecha tal cual, pero si quieres hacer cálculos 
                    // de horas, hazlos en el frontend para evitar confusiones.
                }
            });
        }

        return res.status(200).json(graphData);
    } catch (error) {
        return res.status(500).json({ error: error.message });
    }
}
