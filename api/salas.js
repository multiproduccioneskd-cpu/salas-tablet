// api/salas.js - Backend que conecta con SharePoint

export default async function handler(req, res) {
    // Configuración para evitar que el navegador o Vercel guarden caché vieja
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

    try {
        // AQUÍ ESTÁ EL TRUCO:
        // Agregamos el parámetro ?$orderby=ID desc al final de la URL.
        // Esto obliga a SharePoint a enviarte los eventos MÁS NUEVOS primero.
        // Si SharePoint tiene un límite de respuesta, cortará los antiguos (mayo) 
        // y dejará pasar tus pruebas de la tarde (junio).
        
        const sharepointUrl = "https://graph.microsoft.com/v1.0/sites/TU_SITE_ID/lists/TU_LIST_ID/items?expand=fields&$orderby=ID desc&$top=100";
        
        const response = await fetch(sharepointUrl, {
            headers: {
                'Authorization': `Bearer ${process.env.MICROSOFT_GRAPH_TOKEN}`,
                'Accept': 'application/json'
            }
        });

        const data = await response.json();
        
        // Enviamos la data al HTML
        return res.status(200).json(data);

    } catch (error) {
        return res.status(500).json({ error: "Error conectando a SharePoint" });
    }
}
