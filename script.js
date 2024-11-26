
async function analyzeFiles() {
    const file1 = document.getElementById("file1").files[0];
    const file2 = document.getElementById("file2").files[0];
    const file3 = document.getElementById("file3").files[0];
    if (!file1 || !file2 || !file3) {
        alert("Por favor, carga los 3 archivos.");
        return;
    }

    const data1 = await readExcelFile(file1);
    const data2 = await readExcelFile(file2);
    const data3 = await readExcelFile(file3);

    // Ajuste para asegurarnos de que encontramos la columna de estado
    const summary1 = countStates(data1, ["Estado Helix", "estado helix"]);
    const summary2 = countStates(data2, ["Estado", "estado"]);

    // Mostrar el resumen en la tabla
    displaySummary(summary1, summary2);

    // Comparar los IDs y mostrar los que están en OnNet pero no en W45
    compareIds(data1, data2);

    // Comparar los IDs y mostrar los que están en OnNet pero no en Entel, y verificar en el tercer archivo
    additionalComparison(data1, data2, data3);

    // Mostrar los casos con id de correlacion vacío
    displayEmptyCorrelation(data2);

    // Mostrar las diferencias en los estados "CANCELADO"
    displayCancelledDifferences(data1, data2);
}

function formatExcelDate(excelDate) {
    if (!excelDate || isNaN(excelDate)) return "No encontrado";
        const date = new Date((excelDate - 25569) * 86400 * 1000); // Conversión desde el formato de fecha Excel
        return date.toISOString().split("T")[0]; // Formato: YYYY-MM-DD
}

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            resolve(json);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function countStates(data, possibleColumns) {
    const headers = data[0];
    let stateIndex = -1;

    // Buscar el índice de la columna que coincida con los posibles nombres de columna
    for (let col of possibleColumns) {
        stateIndex = headers.findIndex(header => header && header.toLowerCase() === col.toLowerCase());
        if (stateIndex !== -1) break;
    }

    // Si no se encuentra ninguna columna de estado, devolver un objeto vacío
    if (stateIndex === -1) return {};

    // Inicializar el objeto de conteo para los estados específicos
    const stateCounts = {
        "RESUELTO": 0,
        "CANCELADO": 0,
        "CERRADO": 0,
        "EN CURSO": 0,
        "SIN INCIDENTES": 0,
        "NO REPORTADO POR ONNET": 0,
        "DIFERENCIAS FECHA CIERRE": 0
    };

    // Contar las ocurrencias de cada estado en la columna identificada
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const state = row[stateIndex];
        if (state === "Cerrado" || state === "Cerrada") {
            stateCounts["CERRADO"] += 1;
        } else if (state === "Cancelado" || state === "Cancelada") {
            stateCounts["CANCELADO"] += 1;
        } else if (stateCounts.hasOwnProperty(state)) {
            stateCounts[state] += 1;
        }
    }
    return stateCounts;
}

function displaySummary(summary1, summary2) {
    const tbody = document.getElementById("tableBody");
    tbody.innerHTML = "";

    const states = ["RESUELTO", "CERRADO", "CANCELADO", "EN CURSO", "SIN INCIDENTES", "NO REPORTADO POR ONNET", "DIFERENCIAS FECHA CIERRE"];
    let totalEntel = 0, totalOnNet = 0;

    states.forEach(state => {
        const count1 = summary1[state] || 0;
        const count2 = summary2[state] || 0;
        totalEntel += count1;
        totalOnNet += count2;

        const row = `<tr>
            <td>${state}</td>
            <td>${count1}</td>
            <td>${count2}</td>
        </tr>`;
        tbody.innerHTML += row;
    });

    // Mostrar los totales en el pie de la tabla
    document.getElementById("totalEntel").innerText = totalEntel;
    document.getElementById("totalOnNet").innerText = totalOnNet;
}

function compareIds(data1, data2) {
    const headers1 = data1[0];
    const headers2 = data2[0];

    // Buscar índices de las columnas de interés
    const caseIndexEntel = headers1.findIndex(header => header && header.toLowerCase() === "case");
    const numeroIndexOnNet = headers2.findIndex(header => header && header.toLowerCase() === "número");
    const estadoHelixIndexEntel = headers1.findIndex(header => header && header.toLowerCase() === "estado helix");

    if (caseIndexEntel === -1 || numeroIndexOnNet === -1 || estadoHelixIndexEntel === -1) {
        alert("No se encontraron las columnas para la tabla no reportados onnet.");
        return;
    }

    // Crear un Set de los números en OnNet para comparación rápida
    const casesInOnNet = new Set(data2.slice(1).map(row => row[numeroIndexOnNet]));
    const comparisonBody = document.getElementById("comparisonBody");
    comparisonBody.innerHTML = "";

    // Contador de casos no reportados por OnNet
    let notReportedByOnNetCount = 0;

    // Revisar los casos en Entel y mostrar los que OnNet no tiene
    for (let i = 1; i < data1.length; i++) {
        const row = data1[i];
        const caseIdEntel = row[caseIndexEntel];
        const estadoHelix = row[estadoHelixIndexEntel] || "";

        if (caseIdEntel && !casesInOnNet.has(caseIdEntel)) {
            notReportedByOnNetCount++; // Aumenta el contador
            const rowHTML = `<tr>
                <td>${caseIdEntel}</td>
                <td>${estadoHelix}</td>
                <td>NO INFORMA</td>
            </tr>`;
            comparisonBody.innerHTML += rowHTML;
        }
    }
    updateNoReportedByOnNet(notReportedByOnNetCount);
}
function updateNoReportedByOnNet(count) {
    // Buscar el elemento correspondiente en la tabla de resumen
    const noReportedCell = document.querySelector("#tableBody tr:nth-child(6) td:nth-child(3)"); // Fila 6, Columna OnNet
    const totalOnNetCell = document.getElementById("totalOnNet");

    // Convertir los valores a enteros y sumar el conteo de "NO REPORTADO POR ONNET"
    const currentNoReportedValue = parseInt(noReportedCell.innerText) || 0;
    const newNoReportedValue = currentNoReportedValue + count;
    noReportedCell.innerText = newNoReportedValue;

    // Actualizar el total en la columna "OnNet"
    const currentTotalOnNet = parseInt(totalOnNetCell.innerText) || 0;
    const newTotalOnNet = currentTotalOnNet + count;
    totalOnNetCell.innerText = newTotalOnNet;
}

function additionalComparison(data1, data2, data3) {
    const headers1 = data1[0];
    const headers2 = data2[0];
    const headers3 = data3[0];

    const caseIndexEntel = headers1.findIndex(header => header && header.toLowerCase() === "case");
    const numeroIndexOnNet = headers2.findIndex(header => header && header.toLowerCase() === "número");
    const estadoIndexOnNet = headers2.findIndex(header => header && header.toLowerCase() === "estado");
    const fechaOnnet = headers2.findIndex(header => header && header.toLowerCase() === "cerrado");
    const caseIndexNewDoc = headers3.findIndex(header => header && header.toLowerCase() === "case");
    const estadoHelixIndexNewDoc = headers3.findIndex(header => header && header.toLowerCase() === "estado helix");
    const fechaHelixNewDoc = headers3.findIndex(header => header && header.toLowerCase() === "fecha de cierre helix");

    if (caseIndexEntel === -1 || numeroIndexOnNet === -1 || estadoIndexOnNet === -1 || fechaOnnet === -1 || caseIndexNewDoc === -1 || estadoHelixIndexNewDoc === -1 || fechaHelixNewDoc === -1) {
        alert("No se encontraron las columnas necesarias.");
        return;
    }

    const casesInEntel = new Set(data1.slice(1).map(row => row[caseIndexEntel]));
    const ongoingBody = document.getElementById("additionalComparisonBody");
    const nonOngoingBody = document.getElementById("nonOngoingEntelBody");

    let additionalEnCurso = 0;
    let additionalNoEnCurso = 0;

    ongoingBody.innerHTML = "";
    nonOngoingBody.innerHTML = "";

    for (let i = 1; i < data2.length; i++) {
        const rowOnNet = data2[i];
        const numeroOnNet = rowOnNet[numeroIndexOnNet];
        const estadoOnNet = rowOnNet[estadoIndexOnNet];
        const fechaCierreOnNet = formatExcelDate(rowOnNet[fechaOnnet]);

        if (numeroOnNet && !casesInEntel.has(numeroOnNet)) {
            const matchingRowInNewDoc = data3.find(row => row[caseIndexNewDoc] === numeroOnNet);
            const estadoHelix = matchingRowInNewDoc ? matchingRowInNewDoc[estadoHelixIndexNewDoc] : "No encontrado";
            const fechaHelix = matchingRowInNewDoc ? formatExcelDate(matchingRowInNewDoc[fechaHelixNewDoc]) : "No encontrado";

            if (estadoHelix === "EN CURSO") {
                additionalEnCurso += 1;
            }else if(estadoHelix != "EN CURSO" && fechaHelix != "No encontrado"){
                additionalNoEnCurso += 1;
            }

            const rowHTML = `<tr>
                <td>${numeroOnNet}</td>
                <td>${estadoHelix}</td>
                <td>${estadoOnNet}</td>
                </tr>`;

            const rowHTML2 = `<tr>
                <td>${numeroOnNet}</td>
                <td>${estadoHelix}</td>
                <td>${estadoOnNet}</td>
                <td>${fechaHelix}</td>
                <td>${fechaCierreOnNet}</td>
                </tr>`;

            if (estadoHelix === "EN CURSO") {
                ongoingBody.innerHTML += rowHTML;
            } else {
                nonOngoingBody.innerHTML += rowHTML2;
            }
        }
    }
    // Actualizar el resumen en la primera tabla
    document.getElementById("summaryTable").querySelectorAll("tbody tr").forEach(row => {
        const state = row.cells[0].innerText;
        if (state === "EN CURSO") {
            row.cells[1].innerText = parseInt(row.cells[1].innerText) + additionalEnCurso;
        }
        if (state === "DIFERENCIAS FECHA CIERRE") {
            row.cells[1].innerText = parseInt(row.cells[1].innerText) + additionalNoEnCurso;
        }
    });
    // Actualizar el total en el pie de la tabla
    let newTotalEntel = 0;
    document.getElementById("summaryTable").querySelectorAll("tbody tr").forEach(row => {
        const count = parseInt(row.cells[1].innerText) || 0;
        newTotalEntel += count;
    });
    document.getElementById("totalEntel").innerText = newTotalEntel;
}

function displayEmptyCorrelation(data2) {
    const headers2 = data2[0];

    // Buscar índice de las columnas necesarias
    const numeroIndexOnNet = headers2.findIndex(header => header && header.toLowerCase() === "número");
    const estadoIndexOnNet = headers2.findIndex(header => header && header.toLowerCase() === "estado");
    const idCorrelacionIndex = headers2.findIndex(header => header && header.toLowerCase() === "id de correlación");

    if (numeroIndexOnNet === -1 || estadoIndexOnNet === -1 || idCorrelacionIndex === -1) {
        alert("No se encontraron las columnas necesarias en el archivo de ONNET.");
        return;
    }

    // Filtrar las filas donde la columna "id de correlacion" esté vacía
    const emptyCorrelationBody = document.getElementById("emptyCorrelationBody");
    emptyCorrelationBody.innerHTML = "";

    let noCorrelationCount = 0; // Contador para los casos sin id

    data2.slice(1).forEach(row => {
        const idCorrelacion = row[idCorrelacionIndex];

        if (!idCorrelacion) {  // Si el "id de correlacion" está vacío o no tiene información
            const numeroOnNet = row[numeroIndexOnNet];
            const estadoOnNet = row[estadoIndexOnNet];

            const rowHTML = `<tr>
                <td>${numeroOnNet}</td>
                <td>SIN INCIDENTE</td>
                <td>${estadoOnNet}</td>
            </tr>`;
            emptyCorrelationBody.innerHTML += rowHTML;
            noCorrelationCount++;
        }
    });
    // Actualizar la tabla resumen
    updateNoINC(noCorrelationCount);
}

function updateNoINC(count) {
    // Buscar el elemento correspondiente en la tabla de resumen
    const noIncCell = document.querySelector("#tableBody tr:nth-child(5) td:nth-child(2)"); // Fila 5, Columna Entel
    const totalEntelCell = document.getElementById("totalEntel");

    // Convertir los valores a enteros y sumar el conteo de "NO REPORTADO POR ONNET"
    const currentNoReportedValue = parseInt(noIncCell.innerText) || 0;
    const newNoReportedValue = currentNoReportedValue + count;
    noIncCell.innerText = newNoReportedValue;

    // Actualizar el total en la columna "OnNet"
    const currentTotalEntel = parseInt(totalEntelCell.innerText) || 0;
    const newTotalEntel = currentTotalEntel + count;
    totalEntelCell.innerText = newTotalEntel;
}

function displayCancelledDifferences(data1, data2) {
    const headers1 = data1[0]; // Encabezados del primer archivo (Entel)
    const headers2 = data2[0]; // Encabezados del segundo archivo (OnNet)

    // Buscar índices de las columnas necesarias
    const caseIndexEntel = headers1.findIndex(header => header && header.toLowerCase() === "case");
    const estadoHelixIndexEntel = headers1.findIndex(header => header && header.toLowerCase() === "estado helix");
    const numeroIndexOnNet = headers2.findIndex(header => header && header.toLowerCase() === "número");
    const estadoIndexOnNet = headers2.findIndex(header => header && header.toLowerCase() === "estado");

    if (caseIndexEntel === -1 || estadoHelixIndexEntel === -1 || numeroIndexOnNet === -1 || estadoIndexOnNet === -1) {
        alert("No se encontraron las columnas necesarias en los archivos.");
        return;
    }

    let sumaResuelto = 0;

    // Filtrar las filas donde el estado en Entel es "CANCELADO" y el estado en OnNet es diferente
    const cancelledDifferencesBody = document.getElementById("cancelledDifferencesBody");
    cancelledDifferencesBody.innerHTML = ""; // Limpiar la tabla antes de llenarla

    data1.slice(1).forEach(rowEntel => {
        const caseEntel = rowEntel[caseIndexEntel];
        const estadoHelixEntel = rowEntel[estadoHelixIndexEntel];

        // Verificar si el estado en Entel es "CANCELADO"
        if (estadoHelixEntel && estadoHelixEntel.toUpperCase() === "CANCELADO") {
            // Buscar el "número" correspondiente en el segundo archivo (OnNet)
            const matchingRowInOnNet = data2.find(rowOnNet => rowOnNet[numeroIndexOnNet] === caseEntel);

            if (matchingRowInOnNet) {
                const estadoOnNet = matchingRowInOnNet[estadoIndexOnNet];

                // Verificar si el estado en OnNet es diferente al de Entel
                if (estadoHelixEntel.toUpperCase() !== estadoOnNet.toUpperCase()) {
                    const rowHTML = `<tr>
                <td>${caseEntel}</td>
                <td>${estadoHelixEntel}</td>
                <td>${estadoOnNet}</td>
            </tr>`;
                    cancelledDifferencesBody.innerHTML += rowHTML;
                    sumaResuelto++;
                }
            }
        }
    });
    updateResuelto(sumaResuelto);
}

function updateResuelto(count) {
    // Buscar el elemento correspondiente en la tabla de resumen
    const ResueltoCell = document.querySelector("#tableBody tr:nth-child(1) td:nth-child(3)");
    const totalOnnetCell = document.getElementById("totalOnNet");

    // Convertir los valores a enteros y sumar el conteo de "NO REPORTADO POR ONNET"
    const currentNoReportedValue = parseInt(ResueltoCell.innerText) || 0;
    const newNoReportedValue = currentNoReportedValue + count;
    ResueltoCell.innerText = newNoReportedValue;

    // Actualizar el total en la columna "OnNet"
    const currentTotalEntel = parseInt(totalOnnetCell.innerText) || 0;
    const newTotalEntel = currentTotalEntel + count;
    totalOnnetCell.innerText = newTotalEntel;
}

function downloadExcel() {
    // Crear un libro de Excel
    const wb = XLSX.utils.book_new();

    // Crear una hoja para las tablas generadas en HTML
    const htmlTablesSheet = [];

    // Capturar todas las tablas generadas por el HTML
    const tables = document.querySelectorAll('table');

    tables.forEach((table, index) => {
        const sheet = XLSX.utils.table_to_sheet(table);

        // Convertir cada tabla a un array bidimensional y agregarla a htmlTablesSheet
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // if (index === 0) {
        //     // Agregar encabezados personalizados al inicio
        //     htmlTablesSheet.push(["CASOS", "ESTADO ENTEL", "ESTADO ONNET"]);
        // }

        // Añadir las filas de la tabla
        htmlTablesSheet.push(...rows);

        // Agregar espacio entre tablas para separación visual
        htmlTablesSheet.push([]);
    });

    // Crear la hoja de Excel con el contenido generado
    const htmlTablesExcelSheet = XLSX.utils.aoa_to_sheet(htmlTablesSheet);

    // Aplicar formato al encabezado de la hoja generada
    const headerRange = XLSX.utils.decode_range(htmlTablesExcelSheet["!ref"]);
    for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        if (!htmlTablesExcelSheet[cellAddress]) continue;

        htmlTablesExcelSheet[cellAddress].s = {
            fill: { fgColor: { rgb: "FFFF00" } }, // Fondo amarillo
            font: { bold: true }                 // Texto en negrita
        };
    }

    XLSX.utils.book_append_sheet(wb, htmlTablesExcelSheet, "Resumen");

    // Agregar las hojas de los archivos originales (Entel y OnNet)
    const file1 = document.getElementById("file1").files[0];
    const file2 = document.getElementById("file2").files[0];

    if (file1 && file2) {
        Promise.all([readExcelFile(file1), readExcelFile(file2)]).then(([data1, data2]) => {
            const sheet1 = XLSX.utils.aoa_to_sheet(data1);
            const sheet2 = XLSX.utils.aoa_to_sheet(data2);

            XLSX.utils.book_append_sheet(wb, sheet1, "Entel");
            XLSX.utils.book_append_sheet(wb, sheet2, "OnNet");

            // Descargar el archivo Excel
            XLSX.writeFile(wb, "Conciliacion_W.xlsx");
        }).catch((error) => {
            alert("Error al procesar los archivos: " + error);
        });
    } else {
        XLSX.writeFile(wb, "Conciliacion_W.xlsx");
    }
}