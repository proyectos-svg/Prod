// Agregamos una variable para rastrear el serial actual procesándose
let serialActivo = ""; 

// 2. LÓGICA DE ESCÁNER (Optimizado para no saturar la cámara)
async function toggleCamara() {
    if (!html5QrScanner) {
        html5QrScanner = new Html5Qrcode("reader");
        try {
            await html5QrScanner.start(
                { facingMode: "environment" }, 
                { fps: 15, qrbox: 250 }, 
                (text) => { 
                    if(text.length >= 50) {
                        // Evitar procesar el mismo código múltiples veces seguidas
                        if (text === serialActivo) return; 
                        
                        serialActivo = text;
                        autoCargarDatos(text);
                    }
                }
            );
        } catch (err) { alert("Error de cámara: " + err); html5QrScanner = null; }
    } else {
        await html5QrScanner.stop();
        html5QrScanner = null;
        document.getElementById("reader").innerHTML = "";
    }
}

// 3. CARGA EN INPUT (Aquí se cumple tu regla de las 2 cantidades)
function autoCargarDatos(serial) {
    // Extraemos la cantidad sugerida del serial (pos 47-50)
    const cantSugerida = parseInt(serial.substring(47, 50)) || 0;
    
    // La ponemos en el input para que el operador la vea y edite si es necesario
    const inputCant = document.getElementById("cantidadRealizada");
    inputCant.value = cantSugerida;
    
    // Feedback visual de que se leyó algo
    inputCant.style.backgroundColor = "#e8f5e9"; 
    setTimeout(() => { inputCant.style.backgroundColor = "white"; }, 1000);
    
    mostrarFeedback();
}

// 4. PROCESAMIENTO DE DATOS (Parsing limpio)
function parsearSerial(s, maquina, origen, cantFinal) {
    const pesoUnitario = parseFloat(s.substring(64, 71)) || 0;
    const cantidad = parseInt(cantFinal) || 0;

    return {
        "id": Date.now(),
        "Maquina": maquina,
        "Origen": origen,
        "Cant_Realizada": cantidad, // Cantidad que viene del INPUT (ya editada)
        "WorkOrder": s.substring(0, 5).trim(),
        "Core_Number": s.substring(5, 18).trim(),
        "Acero": s.substring(41, 45).trim(), // Corregido según tu prompt original (41-45)
        "Weight_Normal": pesoUnitario,
        "Peso_Total": (pesoUnitario * cantidad).toFixed(2),
        "Fecha": new Date().toLocaleString()
    };
}

function procesarGuardado() {
    const modo = document.getElementById("filtroMaquina").value;
    const serialManual = document.getElementById("serialTexto") ? document.getElementById("serialTexto").value.trim() : "";
    
    // Usamos el serial activo (del QR) o el manual
    const serial = serialActivo || serialManual;
    const cantidadDelInput = document.getElementById("cantidadRealizada").value;

    if (serial.length < 50) return alert("Error: Debe escanear o ingresar un serial válido.");
    if (modo === "TODAS") return alert("Selecciona una Máquina (M6, M7 o M8).");

    const obj = parsearSerial(serial, modo, document.getElementById("origenDato").value, cantidadDelInput);

    if (modo === "CONTAR") {
        registrosTemporales.unshift(obj);
    } else {
        registrosPermanentes.unshift(obj);
        localStorage.setItem('datos_prod_v6', JSON.stringify(registrosPermanentes));
    }
    
    // IMPORTANTE: Resetear serialActivo para permitir volver a escanear el mismo si fuera necesario
    serialActivo = ""; 
    limpiarInputs();
    actualizarVista();
}
