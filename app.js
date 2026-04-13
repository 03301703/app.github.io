/**
 * Flores El Trigal - App Perdidas v4 (PWA)
 * Core Logic Module
 */

window.app = (function() {
    'use strict';

    // -- State --
    let inventory = [];
    let currentBedCauses = [];
    let savedRecords = [];

    // -- Selectors --
    const els = {
        fileInput: document.getElementById('fileInput'),
        selBloque: document.getElementById('selBloque'),
        selCama: document.getElementById('selCama'),
        infoVariedad: document.getElementById('infoVariedad'),
        v_nombre: document.getElementById('v_nombre'),
        v_plantas: document.getElementById('v_plantas'),
        causaInput: document.getElementById('causaInput'),
        cantidadInput: document.getElementById('cantidadInput'),
        otraCausaDiv: document.getElementById('otraCausaDiv'),
        nuevaCausaNombre: document.getElementById('nuevaCausaNombre'),
        sumaTallos: document.getElementById('sumaTallos'),
        porcTotal: document.getElementById('porcTotal'),
        cuerpoTabla: document.getElementById('cuerpoTabla'),
        totalGralTallos: document.getElementById('totalGralTallos'),
        statusCarga: document.getElementById('statusCarga'),
        loader: document.getElementById('loader')
    };

    // -- Initialization --
    function init() {
        const saved = localStorage.getItem('trigal_perdidas_v4');
        if (saved) {
            try {
                savedRecords = JSON.parse(saved);
                renderHistory();
            } catch (e) {
                console.error("Error loading saved data", e);
            }
        }
    }

    // -- Persistence --
    function saveData() {
        localStorage.setItem('trigal_perdidas_v4', JSON.stringify(savedRecords));
    }

    function limpiarCache() {
        if (confirm("¿Estás seguro de que quieres borrar todos los registros de hoy?")) {
            savedRecords = [];
            localStorage.removeItem('trigal_perdidas_v4');
            renderHistory();
            location.reload();
        }
    }

    // -- Utils --
    function normalize(text) {
        if (!text) return "";
        return String(text).toUpperCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .replace(/Ñ/g, "N")
            .trim();
    }

    function showLoader(show) {
        els.loader.style.display = show ? 'flex' : 'none';
    }

    // -- Excel Processing --
    function procesarExcel() {
        const file = els.fileInput.files[0];
        if (!file) {
            alert("Por favor, selecciona un archivo primero.");
            return;
        }

        showLoader(true);
        const reader = new FileReader();
        
        reader.onerror = () => {
            showLoader(false);
            alert("Error al leer el archivo.");
        };

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(firstSheet);

                inventory = json.map(row => ({
                    bloque: normalize(row.bloque || row.Bloque || ""),
                    cama: String(row.cm || row.cama || row.Cama || ""),
                    variedad: normalize(row.nom_flor || row.variedad || row.Variedad || "N/A"),
                    plantas: parseInt(row.plantas || row.Plantas || row.cantidad || row.Plantas_cama || 0)
                })).filter(item => item.bloque);

                if (inventory.length > 0) {
                    els.statusCarga.classList.add('success');
                    poblarBloques();
                } else {
                    alert("No se encontraron registros válidos en el archivo.");
                }
            } catch (err) {
                console.error(err);
                alert("Error al procesar el Excel. Verifica el formato.");
            } finally {
                showLoader(false);
            }
        };

        reader.readAsArrayBuffer(file);
    }

    function poblarBloques() {
        const bloques = [...new Set(inventory.map(d => d.bloque))].sort();
        els.selBloque.innerHTML = '<option value="">-- Seleccionar Bloque --</option>';
        bloques.forEach(b => {
            els.selBloque.innerHTML += `<option value="${b}">${b}</option>`;
        });
        els.selBloque.disabled = false;
    }

    function cargarCamas() {
        const blq = els.selBloque.value;
        if (!blq) return;

        const camas = inventory.filter(d => d.bloque === blq).map(d => d.cama);
        const uniqueCamas = [...new Set(camas)].sort((a,b) => {
             // Handle numeric sorting for bed numbers
             const numA = parseInt(a);
             const numB = parseInt(b);
             if (isNaN(numA) || isNaN(numB)) return a.localeCompare(b);
             return numA - numB;
        });

        els.selCama.innerHTML = '<option value="">-- Seleccionar --</option>';
        uniqueCamas.forEach(c => {
            els.selCama.innerHTML += `<option value="${c}">${c}</option>`;
        });
        els.selCama.disabled = false;
        els.infoVariedad.style.display = 'none';
    }

    function autocompletarTodo() {
        const blq = els.selBloque.value;
        const cam = els.selCama.value;
        const record = inventory.find(d => d.bloque === blq && d.cama === cam);

        if (record) {
            els.infoVariedad.style.display = 'block';
            els.v_nombre.innerText = record.variedad;
            els.v_plantas.innerText = record.plantas;
            currentBedCauses = [];
            recalcular();
        }
    }

    // -- Loss Logic --
    function verificarOtraCausa() {
        els.otraCausaDiv.style.display = (els.causaInput.value === "OTRA") ? "block" : "none";
    }

    function agregarCausa() {
        let causa = els.causaInput.value;
        const otraNombre = els.nuevaCausaNombre.value.trim();
        const cant = parseInt(els.cantidadInput.value) || 0;
        const totalPlantas = parseInt(els.v_plantas.innerText) || 0;

        if (causa === "OTRA") {
            if (!otraNombre) return alert("Por favor escribe el nombre de la causa.");
            causa = normalize(otraNombre);
        }

        if (!causa || cant <= 0 || totalPlantas === 0) {
            alert("Completa todos los datos antes de añadir.");
            return;
        }

        // Limit check
        const currentSum = currentBedCauses.reduce((acc, item) => acc + item.cantidad, 0);
        if (currentSum + cant > totalPlantas) {
            alert(`Error: La suma de pérdidas (${currentSum + cant}) excede la siembra (${totalPlantas}).`);
            return;
        }

        const existing = currentBedCauses.find(c => c.nombre === causa);
        if (existing) {
            existing.cantidad += cant;
        } else {
            currentBedCauses.push({ nombre: causa, cantidad: cant });
        }

        // Reset inputs
        els.cantidadInput.value = "";
        els.causaInput.value = "";
        els.nuevaCausaNombre.value = "";
        verificarOtraCausa();
        recalcular();
    }

    function recalcular() {
        const totalPlantas = parseInt(els.v_plantas.innerText) || 0;
        const suma = currentBedCauses.reduce((acc, c) => acc + c.cantidad, 0);
        els.sumaTallos.innerText = suma;
        els.porcTotal.innerText = (totalPlantas > 0 ? (suma / totalPlantas * 100).toFixed(2) : "0.00") + "%";
    }

    function guardarFila() {
        const blq = els.selBloque.value;
        const cam = els.selCama.value;
        const plants = els.v_plantas.innerText;
        const prodType = document.querySelector('input[name="prodType"]:checked').value;

        if (!blq || !cam || currentBedCauses.length === 0) {
            alert("No hay datos para guardar.");
            return;
        }

        const maxCausa = [...currentBedCauses].sort((a,b) => b.cantidad - a.cantidad)[0];

        const record = {
            id: Date.now(),
            producto: prodType,
            bloque: blq,
            cama: cam,
            variedad: els.v_nombre.innerText,
            siembra: plants,
            totalBajas: parseInt(els.sumaTallos.innerText),
            porcentaje: els.porcTotal.innerText,
            causasRaw: [...currentBedCauses],
            causasHTML: currentBedCauses.map(c => {
                const isMax = c.nombre === maxCausa.nombre;
                return `<span class="causa-tag ${isMax ? 'main' : ''}">${c.nombre}: ${c.cantidad}</span>`;
            }).join(" ")
        };

        savedRecords.unshift(record); // Add to top
        saveData();
        renderHistory();

        // Reset for next bed
        currentBedCauses = [];
        els.selCama.value = "";
        els.infoVariedad.style.display = 'none';
        recalcular();
        alert("✅ Registro guardado");
    }

    function renderHistory() {
        els.cuerpoTabla.innerHTML = "";
        let globalTotal = 0;

        savedRecords.forEach(r => {
            globalTotal += r.totalBajas;
            const row = `
                <tr>
                    <td><b>${r.producto}</b></td>
                    <td>B${r.bloque} / C${r.cama}</td>
                    <td>${r.variedad}</td>
                    <td>
                        <div style="font-weight:700;">${r.totalBajas}</div>
                        <div style="font-size:0.75rem; margin-top:2px;">${r.causasHTML}</div>
                    </td>
                </tr>
            `;
            els.cuerpoTabla.innerHTML += row;
        });

        els.totalGralTallos.innerText = `${globalTotal} tallos`;
    }

    // -- Export --
    function exportar() {
        if (savedRecords.length === 0) {
            alert("No hay datos para exportar.");
            return;
        }

        const exportData = savedRecords.map(r => ({
            'Producto': r.producto,
            'Bloque': r.bloque,
            'Cama': r.cama,
            'Variedad': r.variedad,
            'Siembra_Original': r.siembra,
            'Total_Bajas': r.totalBajas,
            'Pérdida_%': r.porcentaje,
            'Detalle_Causas': r.causasRaw.map(c => `${c.nombre}:${c.cantidad}`).join(" | ")
        }));

        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Reporte Bajas");
        
        const date = new Date().toISOString().slice(0,10);
        XLSX.writeFile(wb, `Reporte_Bajas_Trigal_${date}.xlsx`);
    }

    // Run init
    init();

    return {
        procesarExcel,
        cargarCamas,
        autocompletarTodo,
        verificarOtraCausa,
        agregarCausa,
        guardarFila,
        exportar,
        limpiarCache
    };

})();
