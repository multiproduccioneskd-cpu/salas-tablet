export default async function handler(req, res) {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    try {
        // URL NATIVA REST DE SHAREPOINT (Más estable que Graph para listas)
        const siteUrl = `https://${process.env.SHAREPOINT_DOMAIN}/sites/${process.env.SHAREPOINT_SITE_NAME}/_api/web/lists/getbytitle('${process.env.LIST_NAME}')/items`;
        
        const response = await fetch(siteUrl, {
            headers: {
                'Authorization': `Bearer ${process.env.SHAREPOINT_TOKEN}`,
                'Accept': 'application/json;odata=verbose'
            }
        });

        const data = await response.json();
        
        // SharePoint OData devuelve la data en 'd.results'
        const results = data.d ? data.d.results : [];
        return res.status(200).json(results);

    } catch (error) {
        console.error("Error Backend:", error);
        return res.status(500).json({ error: "Fallo de conexión" });
    }
}
