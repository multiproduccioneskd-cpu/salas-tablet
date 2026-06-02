// REEMPLAZA TU FUNCIÓN processSharepointData EN EL index.html POR ESTA:

function processSharepointData(items) {
    initSalas();
    console.log("Datos recibidos de la API:", items); // <--- ESTO NOS DIRÁ LA VERDAD

    items.forEach((item, index) => {
        // Intentamos obtener los campos de varias formas posibles
        const fields = item.fields || item; 
        if (!fields) return;

        const columnaSalaRaw = String(fields.Sala || fields.sala || "").toLowerCase().trim();
        const rawInicio = fields.HoraInicio || fields.casillaTiempo || fields.EventDate || fields.Start || "";
        const rawFin = fields.HoraFin || fields.EndDate || fields.End || "";

        let nombreDeLaReunion = parseField(fields.Asunto || fields.Title || fields.title || "Reunión").trim();

        // ELIMINAMOS EL FILTRO DE FECHA AQUÍ PARA QUE TODO LO QUE LLEGUE SE MUESTRE
        if (columnaSalaRaw) {
            let salaKey = "";
            if (columnaSalaRaw.includes("junta") || columnaSalaRaw.includes("directiva")) salaKey = "Sala Junta Directiva";
            else if (columnaSalaRaw.includes("1")) salaKey = "Sala 1";
            else if (columnaSalaRaw.includes("2")) salaKey = "Sala 2";
            else if (columnaSalaRaw.includes("3")) salaKey = "Sala 3";

            if (salaKey && salasMapeadas[salaKey]) {
                salasMapeadas[salaKey].agenda.push({
                    idMeet: `sp-${index}`,
                    title: nombreDeLaReunion,
                    startTime: formatISODateToHHMM(rawInicio),
                    endTime: formatISODateToHHMM(rawFin)
                });
            }
        }
    });

    renderDashboard();
}
