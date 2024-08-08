// ------------ BOTON PARA MIGRAR DATOS DE LLAMADOS ---------------------------

function LimpiarCoordi() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formS = ss.getSheetByName("Coordi");

    var rangoLimpiar = [ "A2:I" ];
    for (let celda of rangoLimpiar) {
        formS.getRange(celda).clearContent();
    }
}

function pasarCasos() {
    var h1 = SpreadsheetApp.getActive().getSheetByName("Planilla de casos");
    var h2 = SpreadsheetApp.getActive().getSheetByName("Coordi");

    if (h1 && h2) {
        var uFila = h1.getLastRow();
        var rango = h1.getRange("A2:ZZ" + uFila).getValues();

        if (rango && rango.length > 0) {
            var mat = [];

            for (var i = 0; i < rango.length; i++) {
                if (rango[ i ]) {
                    // Verificar si la fila está definida
                    var resultadoFecha = obtenerColumnaFechaMasActualCoordi(rango[ i ]);
                    var indiceColumnaFecha = resultadoFecha[ 1 ];

                    var fila = [
                        rango[ i ][ 0 ], // Columna A
                        rango[ i ][ 1 ], // Columna B
                        resultadoFecha[ 0 ], // Fecha más reciente
                    ];

                    // Agregar las tres columnas adyacentes a la derecha de la fecha más reciente
                    for (var j = 0; j < 3; j++) {
                        fila.push(rango[ i ][ indiceColumnaFecha + j + 1 ]);
                    }

                    // Agregar el valor de la columna D de la hoja "Planilla de casos"
                    fila.push(rango[ i ][ 3 ]); 
                    fila.push(rango[ i ][ 5 ]);

                    mat.push(fila);
                } else {
                    console.error("Fila " + i + " está indefinida.");
                }
            }

            h2.getRange(2, 1, mat.length, mat[ 0 ].length).setValues(mat);
        } else {
            console.error("El rango está indefinido o vacío.");
        }
    } else {
        console.error("No se pudo encontrar una de las hojas de cálculo.");
    }
}

function obtenerColumnaFechaMasActualCoordi(fila) {
    var fechaMasReciente = new Date(1960, 0, 1);
    var indiceColumnaFecha = 0;
    var indiceUltimaFecha = 0; // Nuevo índice para almacenar la columna de la última fecha encontrada

    for (var i = 205; i < fila.length; i++) {
        // Iniciamos desde la columna 206
        var fechaActual = new Date(fila[ i ]);

        if (fechaActual >= fechaMasReciente) {
            fechaMasReciente = fechaActual;
            indiceColumnaFecha = i;
            indiceUltimaFecha = i; // Actualizamos el índice de la última fecha encontrada
        }
    }

    // Buscamos si hay fechas repetidas
    for (var j = indiceColumnaFecha + 1; j < fila.length; j++) {
        var fechaActual = new Date(fila[ j ]);

        if (fechaActual.getTime() === fechaMasReciente.getTime()) {
            indiceUltimaFecha = j; // Actualizamos el índice de la última fecha encontrada
        }
    }

    return [ fechaMasReciente, indiceUltimaFecha ];
}

function BorrarVacio() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Coordi"); // Nombre de tu hoja

    var data = sheet.getDataRange().getValues();

    for (var i = 0; i < data.length; i++) {
        if (
            data[ i ][ 2 ] instanceof Date &&
            data[ i ][ 2 ].getTime() === new Date("1/01/1960").getTime()
        ) {
            // Borrar la información de las columnas A, B, C, D, E y F en la fila i
            for (var j = 0; j < 6; j++) {
                // Recorremos las primeras 6 columnas
                sheet.getRange(i + 1, j + 1).clearContent(); // +1 porque los índices de filas y columnas en Google Sheets empiezan desde 1
            }
        }
    }
}

// Boton Traer llamados (Coordi)

function traerLlamados() {
    // Get the current date and time
    var fechaHora = new Date();

    // Get the script time zone
    var timeZone = Session.getScriptTimeZone();

    // Format the date and time with the desired format and time zone
    var marcaDeTiempo = Utilities.formatDate(
        fechaHora,
        timeZone,
        "dd/MM/yyyy HH:mm:ss"
    );

    // Get the "Coordi" sheet
    var hojaCoordi =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordi");

    // Write the formatted date and time to cell J1
    var celdaJ1 = hojaCoordi.getRange("L1");
    celdaJ1.setValue("Última actualización: " + marcaDeTiempo);

    // Limpiar coordenadas
    LimpiarCoordi();

    LimpiarCasillas()
    LimpiarCasillas();
    LimpiarCasillas2();
    LimpiarCasillas3();
    LimpiarCasillas4();
    LimpiarCasillas5();
    LimpiarCasillas6();
    LimpiarCasillas7();
    LimpiarCasillas8();
    LimpiarCasillas9();
    LimpiarCasillas10();

   
    // Pasar casos
    pasarCasos();

    // Borrar vacío
    BorrarVacio();
    SpreadsheetApp.getUi().alert(
        "Hoja de coordinación actualizada correctamente"
    );
}
