// Agrega esto dentro de tu <script> en la función fetchSalasFromAPI
async function fetchSalasFromAPI() {
    const container = document.getElementById('roomsContainer');
    try {
        const response = await fetch(`/api/salas?refresh=${Date.now()}`);
        if (!response.ok) throw new Error("API Falló");
        
        const data = await response.json();
        const items = Array.isArray(data) ? data : (data.value || []);
        
        if (items.length === 0) {
            container.innerHTML = "<div style='color:white; text-align:center; width:100%;'>API conectada pero no hay eventos. Verifica SharePoint.</div>";
        } else {
            processSharepointData(items);
        }
    } catch (error) {
        console.error("API Desconectada:", error);
        container.innerHTML = `
            <div style='color:#dda126; text-align:center; width:100%; padding:20px;'>
                <h3>⚠️ Visor en Modo Manual</h3>
                <p>La API no responde. Inyecta datos desde la consola (F12) con:</p>
                <code style='background:#1b2a4a; padding:10px; display:block; margin-top:10px;'>
                    crearReunionPrueba("Demo UGM", "Sala Junta Directiva", "19:00", "20:30")
                </code>
            </div>`;
    }
}
