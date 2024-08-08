// Función para guardar las derivaciones en la hoja "Derivaciones" según los datos de una hoja "Ope"
function guardarDerivaciones(opeSheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formS = ss.getSheetByName(opeSheetName); // Obtiene la hoja de cálculo específica ("Ope 1", "Ope 2", etc.)
    const dataS = ss.getSheetByName("Derivaciones"); // Hoja de cálculo donde se guardarán las derivaciones

    // Obtiene los valores de las acciones de derivación de la hoja específica ("Ope 1", "Ope 2", etc.)
    const accion1 = formS.getRange("J28").getDisplayValue();
    const accion2 = formS.getRange("J41").getDisplayValue();
    const accion3 = formS.getRange("J45").getDisplayValue();
    const accion4 = formS.getRange("J47").getDisplayValue();
    const accion5 = formS.getRange("J68").getDisplayValue();

    const COLUMNA_DNI = 7; // Índice de columna para el DNI en la hoja "Derivaciones" (columna G)

    // Función para obtener la última fila con datos en una hoja (ignorando la fila de encabezados)
    function GetLastDataRow_(sheet) {
        return sheet.getLastRow() > 1 ? sheet.getLastRow() : 1;
    }

    // Verifica si alguna acción indica una derivación
    if (accion1 == "derivacion" || accion2 == "derivacion" || accion3 == "derivacion" || accion4 == "derivacion" || accion5 == "derivacion") {
        const dni = formS.getRange("B9").getDisplayValue(); // Obtiene el DNI de la hoja específica

        asignarValores1(formS); // Llama a la función para asignar los valores del formulario (primera parte)
        asignarValores2(formS); // Llama a la función para asignar los valores del formulario (segunda parte)

        // Combina los valores obtenidos en un solo array para insertar/actualizar en "Derivaciones"
        const newRowValues = valores1[0].concat(valores2[0]);

        // Obtiene los datos actuales de la hoja "Derivaciones"
        const data = dataS.getDataRange().getValues();
        let foundIndex = -1;

        // Busca el DNI en la hoja de datos
        for (let i = 1; i < data.length; i++) {
            if (data[i][COLUMNA_DNI - 1] == dni) {
                foundIndex = i + 1; // Encuentra el índice correspondiente sumando 1 (porque los índices de getRange empiezan en 1)
                break;
            }
        }

    // Si encontró el DNI, actualiza la fila; si no, agrega una nueva fila al final de los datos
    if (foundIndex !== -1) {
      dataS.getRange(foundIndex, 4, 1, newRowValues.length).setValues([newRowValues]); // Empieza en la columna D en "Derivaciones"
    } else {
      dataS.getRange(GetLastDataRow_(dataS) + 1, 4, 1, newRowValues.length).setValues([newRowValues]); // Agrega después de la última fila con datos
    }
  }
}
