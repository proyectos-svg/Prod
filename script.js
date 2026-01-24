let registrosPermanentes = [];
let registrosTemporales = [];
let html5QrScanner = null;
let ultimoQrEscaneado = "";

// 1. CARGA INICIAL DE MEMORIA
window.onload = () => {
    const memoria = localStorage.getItem('datos_prod_v6');
    if (memoria) {
        registrosPermanentes = JSON.parse(memoria);
        // Asegurar que cada registro tenga un ID único para poder borrarlo
        registrosPermanentes.forEach((r, idx) => { if (!r.id) r.id = Date.now() + idx; });
    }
    actualizarVista();
};

// 2. LÓGICA DE ESCÁNER Y CÁMARA
async function toggleCamara() {
    if (!html5QrScanner) {
        html5QrScanner = new Html5Qrcode("reader");
        try {
            await html5QrScanner.start({ facingMode: "environment" }, { fps: 15, qrbox: 250 }, (text) => { 
                if(text.length > 40) {
                    ultimoQrEscaneado = text;
                    autoCargarDatos(text);
                    // Reinicio rápido para permitir múltiples escaneos
                    setTimeout(() => { if(html5QrScanner) toggleCamara(); }, 500);
                }
            });
        } catch (err) { alert("Error de cámara: " + err); html5QrScanner = null; }
    } else {
        await html5QrScanner.stop();
        html5QrScanner = null;
        document.getElementById("reader").innerHTML = "";
    }
}

function autoCargarDatos(serial) {
    if (serial.length >= 50) {
        const tq = parseInt(serial.substring(47, 50));
        if (!isNaN(tq)) {
            document.getElementById("cantidadRealizada").value = tq; 
            mostrarFeedback();
        }
    }
}

function mostrarFeedback() {
    const status = document.getElementById("statusScan");
    status.style.display = "block";
    setTimeout(() => { status.style.display = "none"; }, 3000);
}

// 3. PROCESAMIENTO DE DATOS
function parsearSerial(s, maquina, origen, cant) {
    const pesoUnitario = parseFloat(s.substring(64, 71)) || 0;
    const cantidad = parseInt(cant) || 0;
    const acero = s.substring(32, 45).trim(); // Posición 33 a 45

    return {
        "id": Date.now(),
        "Maquina": maquina,
        "Origen": origen,
        "Cant_Realizada": cantidad,
        "WorkOrder": s.substring(0, 5).trim(),
        "Core_Number": s.substring(5, 18).trim(),
        "Acero": acero, 
        "Weight_Normal": pesoUnitario,
        "Peso_Total": (pesoUnitario * cantidad).toFixed(2),
        "Fecha": new Date().toLocaleString()
    };
}

function procesarGuardado() {
    const modo = document.getElementById("filtroMaquina").value;
    const serialManual = document.getElementById("serialTexto") ? document.getElementById("serialTexto").value.trim() : "";
    const serial = ultimoQrEscaneado || serialManual;
    const cantidadEditada = document.getElementById("cantidadRealizada").value;

    if (serial.length < 50) return alert("Error: Serial no detectado o inválido.");
    
    const obj = parsearSerial(serial, modo, document.getElementById("origenDato").value, cantidadEditada);

    if (modo === "CONTAR") {
        registrosTemporales.unshift(obj);
    } else {
        if(modo === "TODAS") return alert("Selecciona una Máquina específica.");
        registrosPermanentes.unshift(obj);
        localStorage.setItem('datos_prod_v6', JSON.stringify(registrosPermanentes));
    }
    
    limpiarInputs();
    actualizarVista();
}

// 4. EXPORTACIÓN A EXCEL (Versión Limpia y Centrada)
function descargarExcelFormatoImagen() {
    if (registrosPermanentes.length === 0) return alert("No hay datos para exportar.");

    let ws_data = [];
    let merges = [];
    
    // Título Principal
    ws_data.push(["REPORTE DE PRODUCCION SUMWIC"]);
    merges.push({ s: {r: 0, c: 0}, e: {r: 0, c: 6} });
    ws_data.push([]); 

    // Fecha
    ws_data.push(["", "", "", "", "", "FECHA:", new Date().toLocaleDateString()]);
    ws_data.push([]); 

    const maquinasMap = [
        { sys: "Máquina 6", label: "M6" },
        { sys: "Máquina 7", label: "M7" },
        { sys: "Máquina 8", label: "M8" }
    ];

    maquinasMap.forEach(maq => {
        const datosMaq = registrosPermanentes.filter(r => r.Maquina === maq.sys);
        if (datosMaq.length === 0) return;

        ws_data.push([maq.label, "", "", "", "", "", ""]);
        let rowMaq = ws_data.length - 1;
        merges.push({ s: {r: rowMaq, c: 0}, e: {r: rowMaq, c: 6} });

        ws_data.push(["DESTINO", "No. PARTE", "W.O.", "ACERO", "CANT.", "PESO", "PESO TOTAL"]);

        let totalPz = 0, totalPeso = 0;

        datosMaq.forEach(r => {
            const cant = parseInt(r.Cant_Realizada) || 0;
            const pesoT = parseFloat(r.Peso_Total) || 0;
            ws_data.push([
                r.Origen,                     
                "0000-" + r.Core_Number,      
                r.WorkOrder,
                r.Acero || "", 
                cant,                         
                parseFloat(r.Weight_Normal),  
                pesoT                         
            ]);
            totalPz += cant; 
            totalPeso += pesoT;
        });

        // Fila de Totales: Solo valores y etiquetas, sin recuadros de tabla
        ws_data.push(["", "", "TOTAL PZ:", totalPz, "TOTAL LBS:", totalPeso.toFixed(2), ""]);
        ws_data.push([]); 
    });

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    ws['!merges'] = merges;

    // APLICACIÓN DE ESTILOS FILTRADOS
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell_ref = XLSX.utils.encode_cell({c: C, r: R});
            const cell = ws[cell_ref];

            // 1. Ignorar celdas vacías (Sin bordes)
            if (!cell || cell.v === undefined || cell.v === "") continue;

            // 2. Estilo base: Todo Centrado
            cell.s = {
                alignment: { horizontal: "center", vertical: "center" },
                font: { name: "Arial", sz: 10 }
            };

            const valor = String(cell.v);
            const esEtiquetaTotal = valor.includes("TOTAL PZ:") || valor.includes("TOTAL LBS:");
            const esTituloPrincipal = R === 0;

            // 3. Aplicar bordes SOLO a la tabla de datos (Excluye títulos y totales)
            if (!esEtiquetaTotal && !esTituloPrincipal) {
                cell.s.border = {
                    top: { style: "thin" }, bottom: { style: "thin" },
                    left: { style: "thin" }, right: { style: "thin" }
                };
            }

            // 4. Negritas estratégicas
            if (R === 0 || esEtiquetaTotal || ["DESTINO", "M6", "M7", "M8"].includes(valor)) {
                cell.s.font.bold = true;
            }
        }
    }

    ws['!cols'] = [{ wch: 12 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 8 }, { wch: 10 }, { wch: 12 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte");
    XLSX.writeFile(wb, `Reporte_SUMWIC_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// 5. UTILIDADES DE VISTA
function actualizarVista() {
    const modo = document.getElementById("filtroMaquina").value;
    const container = document.getElementById("contenedorPrincipal");
    container.innerHTML = "";
    
    if (modo === "CONTAR") {
        if (registrosTemporales.length === 0) { container.innerHTML = "<p style='text-align:center;'>Escanee para contar</p>"; return; }
        const totalLbs = registrosTemporales.reduce((a, b) => a + parseFloat(b.Peso_Total), 0).toFixed(2);
        container.innerHTML = `<div class="weight-card-top"><label>TOTAL ACUMULADO</label><span>${totalLbs} Lbs</span></div>`;
        renderLista(container, registrosTemporales);
    } else {
        const datos = modo === "TODAS" ? registrosPermanentes : registrosPermanentes.filter(r => r.Maquina === modo);
        renderLista(container, datos);
    }
}

function renderLista(container, datos) {
    if (datos.length === 0) { container.innerHTML = "<p style='text-align:center;'>Sin registros</p>"; return; }
    datos.forEach(r => {
        const div = document.createElement("div");
        div.className = "inv-item open";
        div.innerHTML = `
            <div class="detail-row">
                <div class="detail-line"><b>WO: ${r.WorkOrder}</b> <b>${r.Cant_Realizada} pz</b></div>
                <div class="detail-line" style="font-size:0.75rem; color:#666;">
                    <span>Core: ${r.Core_Number}</span>
                    <span>${r.Peso_Total} lbs</span>
                </div>
                <button class="btn-delete-row" onclick="borrarRegistro(${r.id})">✕</button>
            </div>
        `;
        container.appendChild(div);
    });
}

function limpiarInputs() {
    ultimoQrEscaneado = "";
    if(document.getElementById("serialTexto")) document.getElementById("serialTexto").value = "";
    document.getElementById("cantidadRealizada").value = 0;
}

function borrarRegistro(id) {
    if(!confirm("¿Eliminar registro?")) return;
    registrosPermanentes = registrosPermanentes.filter(r => r.id !== id);
    localStorage.setItem('datos_prod_v6', JSON.stringify(registrosPermanentes));
    registrosTemporales = registrosTemporales.filter(r => r.id !== id);
    actualizarVista();
}

function toggleManual() {
    const area = document.getElementById("inputManualArea");
    area.style.display = area.style.display === "none" ? "block" : "none";
}
