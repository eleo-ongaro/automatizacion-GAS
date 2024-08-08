// ---------------------------- OPE 1 -----------------------------------------

// Obtener la hoja de cálculo activa
var hojaCalculo = SpreadsheetApp.getActiveSpreadsheet();
// Obtener la hoja "Ope 1"
var ope1 = hojaCalculo.getSheetByName("Ope 1");

var PlanillaDeCasos = hojaCalculo.getSheetByName("Planilla de casos");

// Función para buscar la información de la última fecha
function buscarInformacionUltimaFecha() {
    var dniCeldaC3 = ope1.getRange("B3").getValue();
    var datos = PlanillaDeCasos.getDataRange().getValues();

    var ultimaFecha = new Date(1900, 0, 1);
    var informacionUltimaFecha = [];

    for (var i = 0; i < datos.length; i++) {
        var fila = datos[ i ];

        if (fila[ 3 ] == dniCeldaC3) {
            var informacionFila = obtenerColumnaFechaMasActual(fila.slice(0));

            if (informacionFila[ 0 ] > ultimaFecha) {
                ultimaFecha = informacionFila[ 0 ];
                informacionUltimaFecha = informacionFila;
            }
        }
    }

    // Verifica si se encontraron fechas antes de devolver la información
    if (informacionUltimaFecha.length > 0) {
        return informacionUltimaFecha;
    } else {
        return []; // Devuelve un array vacío si no se encontraron fechas
    }
}

// Función para obtener la columna de fecha más actual y las 3 columnas siguientes
function obtenerColumnaFechaMasActual(rango) {
    let fechas = [];
    let fechasRepetidas = [];

    // Iterar sobre el rango para obtener todas las fechas
    for (let i = 205; i < rango.length; i++) {
        const fechaActual = new Date(rango[ i ]);
        fechas.push({ fecha: fechaActual, indice: i });
    }

    // Encontrar la fecha más reciente
    const fechaMasReciente = fechas.reduce(
        (max, fecha) => (max.fecha < fecha.fecha ? fecha : max),
        fechas[ 0 ]
    ).fecha;

    // Filtrar las fechas que son iguales a la fecha más reciente
    fechasRepetidas = fechas.filter(
        (fecha) => fecha.fecha.getTime() === fechaMasReciente.getTime()
    );

    if (fechasRepetidas.length > 0) {
        // Ordenar las fechas repetidas por su índice (de izquierda a derecha)
        fechasRepetidas.sort((a, b) => a.indice - b.indice);
        const indiceColumnaFecha =
            fechasRepetidas[ fechasRepetidas.length - 1 ].indice;

        // Obtener las 4 columnas correspondientes a la fecha más a la derecha
        const columnasResultado = rango.slice(
            indiceColumnaFecha,
            indiceColumnaFecha + 4
        );
        return columnasResultado;
    }

    return [];
}

// Función para pegar la información de la última fecha en el "Ope 1"
function pegarInformacionUltimaFecha() {
    var informacionUltimaFecha = buscarInformacionUltimaFecha();

    // Verificar si hay fechas disponibles en la información
    if (informacionUltimaFecha.length > 0 && informacionUltimaFecha[ 0 ]) {
        // Comparar la fecha con la existente en C3
        if (informacionUltimaFecha[ 0 ] != ope1.getRange("B3").getValue()) {
            var rangoDestino = ope1.getRange("E4:H4");
            rangoDestino.setValues([ informacionUltimaFecha ]);
        } else {
            Logger.log("La fecha es la misma, no se realizará la actualización.");
        }
    } else {
        Logger.log("No se encontraron fechas para el DNI especificado.");
    }
}

// Función para buscar la fila en la hoja "Planilla de casos" que coincida con el DNI
function buscarFilaPorDni(dni, hoja) {
    var datos = hoja.getDataRange().getValues();

    for (var i = 0; i < datos.length; i++) {
        if (datos[ i ][ 3 ] == dni) {
            return i + 1;
        }
    }

    return -1;
}

// RANGO DE CELDAS A LIMPIAR (MENOS EL DNI)

var rangoALimpiarMenosDNI = [
    // ULTIMO REGISTRO
    "E4",
    "F4",
    "G4",
    "H4",
    // NUEVO REGISTRO
    "E5",
    //"F5",
    "G5",
    "H5",
    // PROXIMO LLAMADO
    "I5",
    "J5",
    // DATOS PERSONALES
    "B8:C19",

    // OBSERVACIONES
    "D8",
    // Categorizacion
    "B21",
    "D21",
    "F21",

    // Entrevista

    "B23",
    "D23",
    "F23",
    "H23",
    
    "B24",
    "D24",
    "F24",
    "H24",

    "B25",
    "D25",
    "F25",
    "H25",

    "B26",
    "D26",
    "F26",
    "H26",

    "B27",
    "D27",
    "F27",
    "H27",

    "B28",
    "D28",
    "F28",
    "H28",

  //  Convivientes/contacto

    "B30",
    "D30",
    "B32:I41",

  // VACUNACIÓN / SUBSIDIO / SALUD MENTAL

  "B43:B45",
  "D43:D44",
  "F43:F44",
  "H44",

  // GESTACIÓN / PUERPERIO

  "B47",
  "D47",
  "F47",

  // SEGUIMIENTO 1

  "B49:B51",
  "D49:D51",
  "F49:F50",
  "H49:H50",
  "J49:J50",

   // SEGUIMIENTO 2

  "B53:B55",
  "D53:D55",
  "F53:F54",
  "H53:H54",
  "J53:J54",


 // SEGUIMIENTO 3

  "B57:B59",
  "D58:D59",
  "F57:F58",
  "H57:H58",
  "J57:J58",


 // SEGUIMIENTO 4

  "B61:B63",
  "D61:D63",
  "F61:F62",
  "H61:H62",
  "J61:J62",


 // SEGUIMIENTO 5

  "B65:B67",
  "D65:D67",
  "F65:F66",
  "H65:H66",
  "J65:J66",



];

// FUNCION LIMPIAR TODO MENOS EL DNI (para el Boton Buscar)
function LimpiarMenosDNI() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formS = ss.getSheetByName("Ope 1");
    for (var i = 0; i < rangoALimpiarMenosDNI.length; i++) {
        formS.getRange(rangoALimpiarMenosDNI[ i ]).clearContent();
    }
}

// RANGO A LIMPIAR (HOJA OPE COMPLETA)
var rangesToClear = [
    //DNI
    "B3",
      // ULTIMO REGISTRO
    "E4",
    "F4",
    "G4",
    "H4",
    // NUEVO REGISTRO
    "E5",
    //"F5",
    "G5",
    "H5",
    // PROXIMO LLAMADO
    "I5",
    "J5",
    // DATOS PERSONALES
    "B8:C19",

    // OBSERVACIONES
    "D8",
    // Categorizacion
    "B21",
    "D21",
    "F21",

    // Entrevista

    "B23",
    "D23",
    "F23",
    "H23",
    
    "B24",
    "D24",
    "F24",
    "H24",

    "B25",
    "D25",
    "F25",
    "H25",

    "B26",
    "D26",
    "F26",
    "H26",

    "B27",
    "D27",
    "F27",
    "H27",

    "B28",
    "D28",
    "F28",
    "H28",

  //  Convivientes/contacto

    "B30",
    "D30",
    "B32:I41",

  // VACUNACIÓN / SUBSIDIO / SALUD MENTAL

  "B43:B45",
  "D43:D44",
  "F43:F44",
  "H44",

  // GESTACIÓN / PUERPERIO

  "B47",
  "D47",
  "F47",

  // SEGUIMIENTO 1

  "B49:B51",
  "D49:D51",
  "F49:F50",
  "H49:H50",
  "J49:J50",

   // SEGUIMIENTO 2

  "B53:B55",
  "D53:D55",
  "F53:F54",
  "H53:H54",
  "J53:J54",


 // SEGUIMIENTO 3

  "B57:B59",
  "D58:D59",
  "F57:F58",
  "H57:H58",
  "J57:J58",


 // SEGUIMIENTO 4

  "B61:B63",
  "D61:D63",
  "F61:F62",
  "H61:H62",
  "J61:J62",


 // SEGUIMIENTO 5

  "B65:B67",
  "D65:D67",
  "F65:F66",
  "H65:H66",
  "J65:J66",

];

// Función para limpiar las celdas
function Limpiar() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formS = ss.getSheetByName("Ope 1");
    for (var i = 0; i < rangesToClear.length; i++) {
        formS.getRange(rangesToClear[ i ]).clearContent();
    }
}
// Función para crear las nuevas columnas o pegar la información según corresponda
function crearNuevasColumnasOActualizarDatos() {
    // Obtener los datos de las nuevas columnas (G5:J5)
    var datosNuevos = ope1.getRange("E5:H5").getValues();

    // Obtener el DNI del Ope 1 (C3)
    var dniope1 = ope1.getRange("B3").getValue();

    // Buscar la fila en la hoja "Planilla de casos" que coincida con el DNI
    var filaEncontrada = buscarFilaPorDni(dniope1, PlanillaDeCasos);

    // Si se encontró la fila
    if (filaEncontrada !== -1) {
        // Obtener el rango de datos
        var datosPlanillaDeCasos = PlanillaDeCasos.getDataRange().getValues();

        // Encontrar la primera columna vacía después del DNI
        var primeraColumnaVacia = datosPlanillaDeCasos[ filaEncontrada - 1 ].indexOf(
            "",
            205 // es el indice de donde debería aparecer la primer columna nueva con la info de "nuevo registro". En ésta línea es la colimna IE, que es la 69 (como hay q poner ínidce, se empieza x 0, asiq le restamos 1 a 69)
        );

        // Si no se encontró una columna vacía después del DNI, establecerla en la siguiente columna disponible
        if (primeraColumnaVacia === -1) {
            primeraColumnaVacia = datosPlanillaDeCasos[ filaEncontrada - 1 ].length;
        }

        // Determinar el rango donde pegar los datos
        var rangoDestino = PlanillaDeCasos.getRange(
            filaEncontrada,
            primeraColumnaVacia + 1,
            datosNuevos.length,
            datosNuevos[ 0 ].length
        );

        // Pegar los datos en el rango destino
        rangoDestino.setValues(datosNuevos);

        // Agregar encabezados a las nuevas columnas
        var encabezados = [
            "Fecha",
            "Operador",
            "Tipo de llamado",
            "Resultado de llamado",
        ];
        for (var i = 0; i < encabezados.length; i++) {
            PlanillaDeCasos.getRange(1, primeraColumnaVacia + i + 1).setValue(
                encabezados[ i ]
            );
        }

        // Mostrar un mensaje de confirmación
        // Browser.msgBox("Información actualizada correctamente.");
    } else {
        // Mostrar un mensaje de error si no se encontró la fila
        Browser.msgBox(
            "No se encontró la fila para el DNI " + dniope1 + ". Dar aviso a soporte"
        );
    }
}

// Definir una función para establecer los valores en el formulario
function establecerValoresEnFormulario(formulario, fila) {

    // Motivo y fecha de PROXIMO LLAMADO

    formulario.getRange("I5").setValue(fila[ 0 ]);
    formulario.getRange("J5").setValue(fila[ 1 ]);

    // Datos personales

    formulario.getRange("B8").setValue(fila[ 2 ]);
    formulario.getRange("B9").setValue(fila[ 3 ]);
    formulario.getRange("B10").setValue(fila[ 4 ]);
    formulario.getRange("B11").setValue(fila[ 5 ]);
    formulario.getRange("B12").setValue(fila[ 6 ]);
    formulario.getRange("B13").setValue(fila[ 7 ]);
    formulario.getRange("B14").setValue(fila[ 8 ]);
    formulario.getRange("B15").setValue(fila[ 9 ]);
    formulario.getRange("B16").setValue(fila[ 10 ]);
    formulario.getRange("B17").setValue(fila[ 11 ]);
    formulario.getRange("B18").setValue(fila[ 12 ]);
    formulario.getRange("B19").setValue(fila[ 13 ]);

      // OBSERVACIONES

    formulario.getRange("D8").setValue(fila[ 14 ]);


    // Categorizacion

    formulario.getRange("B21").setValue(fila[ 15 ]);
    formulario.getRange("D21").setValue(fila[ 16 ]);
    formulario.getRange("F21").setValue(fila[ 17 ]);

    // ENTREVISTA


    formulario.getRange("B23").setValue(fila[ 18 ]);
    formulario.getRange("D23").setValue(fila[ 19 ]);
    formulario.getRange("F23").setValue(fila[ 20 ]);
    formulario.getRange("H23").setValue(fila[ 21 ]);

    formulario.getRange("B24").setValue(fila[ 22 ]);
    formulario.getRange("D24").setValue(fila[ 23 ]);
    formulario.getRange("F24").setValue(fila[ 24 ]);
    formulario.getRange("H24").setValue(fila[ 25 ]);

    formulario.getRange("B25").setValue(fila[ 26 ]);
    formulario.getRange("D25").setValue(fila[ 27 ]);
    formulario.getRange("F25").setValue(fila[ 28 ]);
    formulario.getRange("H25").setValue(fila[ 29 ]);

    formulario.getRange("B26").setValue(fila[ 30 ]);
    formulario.getRange("D26").setValue(fila[ 31 ]);
    formulario.getRange("F26").setValue(fila[ 32 ]);
    formulario.getRange("H26").setValue(fila[ 33 ]);

    formulario.getRange("B27").setValue(fila[ 34 ]);
    formulario.getRange("D27").setValue(fila[ 35 ]);
    formulario.getRange("F27").setValue(fila[ 36 ]);
    formulario.getRange("H27").setValue(fila[ 37 ]);

    formulario.getRange("B28").setValue(fila[ 38 ]);
    formulario.getRange("D28").setValue(fila[ 39 ]);
    formulario.getRange("F28").setValue(fila[ 40 ]);
    formulario.getRange("H28").setValue(fila[ 41 ]);

    // Accion CETEC

    // formulario.getRange("J28").setValue(fila[ 42 ]);


    // CONVIVIENTES/CONTACTOS

    formulario.getRange("B30").setValue(fila[ 43 ]);
    formulario.getRange("D30").setValue(fila[ 44 ]);

    
    formulario.getRange("B32").setValue(fila[ 45 ]);
    formulario.getRange("C32").setValue(fila[ 46 ]);
    formulario.getRange("D32").setValue(fila[ 47 ]);
    formulario.getRange("E32").setValue(fila[ 48 ]);
    formulario.getRange("F32").setValue(fila[ 49 ]);
    formulario.getRange("G32").setValue(fila[ 50 ]);
    formulario.getRange("H32").setValue(fila[ 51 ]);
    formulario.getRange("I32").setValue(fila[ 52 ]);

    formulario.getRange("B33").setValue(fila[ 53 ]);
    formulario.getRange("C33").setValue(fila[ 54 ]);
    formulario.getRange("D33").setValue(fila[ 55 ]);
    formulario.getRange("E33").setValue(fila[ 56 ]);
    formulario.getRange("F33").setValue(fila[ 57 ]);
    formulario.getRange("G33").setValue(fila[ 58 ]);
    formulario.getRange("H33").setValue(fila[ 59 ]);
    formulario.getRange("I33").setValue(fila[ 60 ]);

    formulario.getRange("B34").setValue(fila[ 61 ]);
    formulario.getRange("C34").setValue(fila[ 62 ]);
    formulario.getRange("D34").setValue(fila[ 63 ]);
    formulario.getRange("E34").setValue(fila[ 64 ]);
    formulario.getRange("F34").setValue(fila[ 65 ]);
    formulario.getRange("G34").setValue(fila[ 66 ]);
    formulario.getRange("H34").setValue(fila[ 67 ]);
    formulario.getRange("I34").setValue(fila[ 68 ]);

    formulario.getRange("B35").setValue(fila[ 69 ]);
    formulario.getRange("C35").setValue(fila[ 70 ]);
    formulario.getRange("D35").setValue(fila[ 71 ]);
    formulario.getRange("E35").setValue(fila[ 72 ]);
    formulario.getRange("F35").setValue(fila[ 73 ]);
    formulario.getRange("G35").setValue(fila[ 74 ]);
    formulario.getRange("H35").setValue(fila[ 75 ]);
    formulario.getRange("I35").setValue(fila[ 76 ]);

    formulario.getRange("B36").setValue(fila[ 77 ]);
    formulario.getRange("C36").setValue(fila[ 78 ]);
    formulario.getRange("D36").setValue(fila[ 79 ]);
    formulario.getRange("E36").setValue(fila[ 80 ]);
    formulario.getRange("F36").setValue(fila[ 81 ]);
    formulario.getRange("G36").setValue(fila[ 82 ]);
    formulario.getRange("H36").setValue(fila[ 83 ]);
    formulario.getRange("I36").setValue(fila[ 84 ]);

    formulario.getRange("B37").setValue(fila[ 85 ]);
    formulario.getRange("C37").setValue(fila[ 86 ]);
    formulario.getRange("D37").setValue(fila[ 87 ]);
    formulario.getRange("E37").setValue(fila[ 88 ]);
    formulario.getRange("F37").setValue(fila[ 89 ]);
    formulario.getRange("G37").setValue(fila[ 90 ]);
    formulario.getRange("H37").setValue(fila[ 91 ]);
    formulario.getRange("I37").setValue(fila[ 92 ]);

    formulario.getRange("B38").setValue(fila[ 93 ]);
    formulario.getRange("C38").setValue(fila[ 94 ]);
    formulario.getRange("D38").setValue(fila[ 95 ]);
    formulario.getRange("E38").setValue(fila[ 96 ]);
    formulario.getRange("F38").setValue(fila[ 97 ]);
    formulario.getRange("G38").setValue(fila[ 98 ]);
    formulario.getRange("H38").setValue(fila[ 99 ]);
    formulario.getRange("I38").setValue(fila[ 100 ]);

    formulario.getRange("B39").setValue(fila[ 101 ]);
    formulario.getRange("C39").setValue(fila[ 102 ]);
    formulario.getRange("D39").setValue(fila[ 103 ]);
    formulario.getRange("E39").setValue(fila[ 104 ]);
    formulario.getRange("F39").setValue(fila[ 105 ]);
    formulario.getRange("G39").setValue(fila[ 106 ]);
    formulario.getRange("H39").setValue(fila[ 107 ]);
    formulario.getRange("I39").setValue(fila[ 108 ]);

    formulario.getRange("B40").setValue(fila[ 109 ]);
    formulario.getRange("C40").setValue(fila[ 110 ]);
    formulario.getRange("D40").setValue(fila[ 111 ]);
    formulario.getRange("E40").setValue(fila[ 112 ]);
    formulario.getRange("F40").setValue(fila[ 113 ]);
    formulario.getRange("G40").setValue(fila[ 114 ]);
    formulario.getRange("H40").setValue(fila[ 115 ]);
    formulario.getRange("I40").setValue(fila[ 116 ]);

    formulario.getRange("B41").setValue(fila[ 117 ]);
    formulario.getRange("C41").setValue(fila[ 118 ]);
    formulario.getRange("D41").setValue(fila[ 119 ]);
    formulario.getRange("E41").setValue(fila[ 120 ]);
    formulario.getRange("F41").setValue(fila[ 121 ]);
    formulario.getRange("G41").setValue(fila[ 122 ]);
    formulario.getRange("H41").setValue(fila[ 123 ]);
    formulario.getRange("I41").setValue(fila[ 124 ]);

    // Accion CETEC

    // formulario.getRange("J41").setValue(fila[ 125 ]);


    // VACUNACIÓN / SUBSIDIO / SALUD MENTAL


    formulario.getRange("B43").setValue(fila[ 126 ]);
    formulario.getRange("D43").setValue(fila[ 127 ]);
    formulario.getRange("F43").setValue(fila[ 128 ]);

    formulario.getRange("B44").setValue(fila[ 129 ]);
    formulario.getRange("D44").setValue(fila[ 130 ]);
    formulario.getRange("F44").setValue(fila[ 131 ]);
    formulario.getRange("H44").setValue(fila[ 132 ]);
    formulario.getRange("B45").setValue(fila[ 133 ]);

    // Accion CETEC

    // formulario.getRange("J45").setValue(fila[ 134 ]);

    // GESTACIÓN / PUERPERIO

    formulario.getRange("B47").setValue(fila[ 135 ]);
    formulario.getRange("D47").setValue(fila[ 136 ]);
    formulario.getRange("F47").setValue(fila[ 137 ]);

    // Accion CETEC

    // formulario.getRange("J47").setValue(fila[ 138 ]);

    // SEGUIMIENTO 1

    formulario.getRange("B49").setValue(fila[ 139 ]);
    formulario.getRange("D49").setValue(fila[ 140 ]);
    formulario.getRange("F49").setValue(fila[ 141 ]);
    formulario.getRange("H49").setValue(fila[ 142 ]);
    formulario.getRange("J49").setValue(fila[ 143 ]);

    formulario.getRange("B50").setValue(fila[ 144 ]);
    formulario.getRange("D50").setValue(fila[ 145 ]);
    formulario.getRange("F50").setValue(fila[ 146 ]);
    formulario.getRange("H50").setValue(fila[ 147 ]);
    formulario.getRange("J50").setValue(fila[ 148 ]);

    formulario.getRange("B51").setValue(fila[ 149 ]);
    formulario.getRange("D51").setValue(fila[ 150 ]);

        // SEGUIMIENTO 2

    formulario.getRange("B53").setValue(fila[ 151 ]);
    formulario.getRange("D53").setValue(fila[ 152 ]);
    formulario.getRange("F53").setValue(fila[ 153 ]);
    formulario.getRange("H53").setValue(fila[ 154 ]);
    formulario.getRange("J53").setValue(fila[ 155 ]);

    formulario.getRange("B54").setValue(fila[ 156 ]);
    formulario.getRange("D54").setValue(fila[ 157 ]);
    formulario.getRange("F54").setValue(fila[ 158 ]);
    formulario.getRange("H54").setValue(fila[ 159 ]);
    formulario.getRange("J54").setValue(fila[ 160 ]);

    formulario.getRange("B55").setValue(fila[ 161 ]);
    formulario.getRange("D55").setValue(fila[ 162 ]);

        // SEGUIMIENTO 3

    formulario.getRange("B57").setValue(fila[ 163 ]);
    formulario.getRange("D57").setValue(fila[ 164 ]);
    formulario.getRange("F57").setValue(fila[ 165 ]);
    formulario.getRange("H57").setValue(fila[ 166 ]);
    formulario.getRange("J57").setValue(fila[ 167 ]);

    formulario.getRange("B58").setValue(fila[ 168 ]);
    formulario.getRange("D58").setValue(fila[ 169 ]);
    formulario.getRange("F58").setValue(fila[ 170 ]);
    formulario.getRange("H58").setValue(fila[ 171 ]);
    formulario.getRange("J58").setValue(fila[ 172 ]);

    formulario.getRange("B59").setValue(fila[ 173 ]);
    formulario.getRange("D59").setValue(fila[ 174 ]);

        // SEGUIMIENTO 4

    formulario.getRange("B61").setValue(fila[ 175 ]);
    formulario.getRange("D61").setValue(fila[ 176 ]);
    formulario.getRange("F61").setValue(fila[ 177 ]);
    formulario.getRange("H61").setValue(fila[ 178 ]);
    formulario.getRange("J61").setValue(fila[ 179 ]);

    formulario.getRange("B62").setValue(fila[ 180 ]);
    formulario.getRange("D62").setValue(fila[ 181 ]);
    formulario.getRange("F62").setValue(fila[ 182 ]);
    formulario.getRange("H62").setValue(fila[ 183 ]);
    formulario.getRange("J62").setValue(fila[ 184 ]);

    formulario.getRange("B63").setValue(fila[ 185 ]);
    formulario.getRange("D63").setValue(fila[ 186 ]);

        // SEGUIMIENTO 5

    formulario.getRange("B65").setValue(fila[ 187 ]);
    formulario.getRange("D65").setValue(fila[ 188 ]);
    formulario.getRange("F65").setValue(fila[ 189 ]);
    formulario.getRange("H65").setValue(fila[ 190 ]);
    formulario.getRange("J65").setValue(fila[ 191 ]);

    formulario.getRange("B66").setValue(fila[ 192 ]);
    formulario.getRange("D66").setValue(fila[ 193 ]);
    formulario.getRange("F66").setValue(fila[ 194 ]);
    formulario.getRange("H66").setValue(fila[ 195 ]);
    formulario.getRange("J66").setValue(fila[ 196 ]);

    formulario.getRange("B67").setValue(fila[ 197 ]);
    formulario.getRange("D67").setValue(fila[ 198 ]);
  

    // Accion CETEC

    // formulario.getRange("J68").setValue(fila[ 199 ]);


}

var NUM_COLUMNA_BUSQUEDA = 3;

// Modificar la función Buscar() para usar la función establecerValoresEnFormulario()
function Buscar() {
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
    var formulario = hojaActiva.getSheetByName("Ope 1"); // Nombre de hoja del formulario
    var valor = formulario.getRange("B3").getDisplayValue();
    var valores = hojaActiva
        .getSheetByName("Planilla de casos")
        .getDataRange()
        .getDisplayValues(); // Nombre de hoja donde se almacenan datos

    for (var i = 0; i < valores.length; i++) {
        var fila = valores[ i ];
        if (fila[ NUM_COLUMNA_BUSQUEDA ] == valor) {
            // Llamar a la función para establecer los valores en el formulario
            establecerValoresEnFormulario(formulario, fila);
            // Salir del bucle después de encontrar el valor
            break;
        }
    }
}

function validarCampos() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var fecha = sheet.getRange("E5").getValue();
    var operador = sheet.getRange("F5").getValue();
    var tipoLlamado = sheet.getRange("G5").getValue();
    var resultadoLlamado = sheet.getRange("H5").getValue();
    var motivo = sheet.getRange("I5").getValue();
    var fechaLlamado = sheet.getRange("J5").getValue();
    var hoy = new Date();
    var dnic3 = sheet.getRange("B3").getValue();
    var dnic9 = sheet.getRange("B9").getValue();

    if (fecha === "") {
        throw new Error("La fecha del nuevo registro, está vacía");
    }

    if (operador === "") {
        throw new Error("El operador está vacío");
    }

    if (tipoLlamado === "") {
        throw new Error("El tipo de llamado está vacío");
    }

    if (resultadoLlamado === "") {
        throw new Error("El resultado de llamado está vacío");
    }

    if (motivo === "") {
        throw new Error("El motivo de próximo llamado está vacío");
    }

    if (fechaLlamado === "" && motivo !== "Fin de los seguimientos" && motivo !== "Tiene alta médica") {
    throw new Error("La fecha de próximo llamado está vacía");
    }

    if (fechaLlamado !== "" && motivo === "Fin de los seguimientos" && motivo === "Tiene alta médica") {
        throw new Error("El fin no debe tener fecha de llamado");
    }

    if (dnic3 === "" && dnic9 === "") {
      throw new Error("Las celdas de DNI estan vacías");
    
    }

    if(dnic3 !== dnic9) {
      throw new Error("Los DNI son diferentes");
    }


    // Convertir fecha de llamado a objeto Date
    var fechaLlamadoDate = new Date(fechaLlamado);
    var fechaDate = new Date(fecha);

    // Establecer las horas, minutos, segundos y milisegundos de la fecha de hoy a cero
    hoy.setHours(0, 0, 0, 0);

    // Establecer las horas, minutos, segundos y milisegundos de la fecha de llamado a cero
    fechaLlamadoDate.setHours(0, 0, 0, 0);
    fechaDate.setHours(0, 0, 0, 0);

    if (fechaLlamado !== "" && fechaLlamadoDate < hoy) {
        throw new Error("La fecha de próximo llamado es anterior a hoy");
    }
    if (fecha !== "" && fechaDate < hoy) {
        throw new Error("La fecha del llamado actual es anterior a hoy");
    }
    if (fecha !== "" && fechaDate > hoy) {
        throw new Error("La fecha del llamado actual es posterior a hoy");
    }
}

// Define la variable valores1 fuera de la función Actualizar
var valores1 = null;

// Función para asignar valores a la variable valores1
function asignarValores1(formulario) {
    valores1 = [
        [

         // Motivo y fecha de PROXIMO LLAMADO

    formulario.getRange("I5").getDisplayValue(),
    formulario.getRange("J5").getDisplayValue(),

    // Datos personales

    formulario.getRange("B8").getDisplayValue(),
    formulario.getRange("B9").getDisplayValue(),
    formulario.getRange("B10").getDisplayValue(),
    formulario.getRange("B11").getDisplayValue(),
    formulario.getRange("B12").getDisplayValue(),
    formulario.getRange("B13").getDisplayValue(),
    formulario.getRange("B14").getDisplayValue(),
    formulario.getRange("B15").getDisplayValue(),
    formulario.getRange("B16").getDisplayValue(),
    formulario.getRange("B17").getDisplayValue(),
    formulario.getRange("B18").getDisplayValue(),
    formulario.getRange("B19").getDisplayValue(),

      
        ],
    ];
}

function asignarValores2(formulario) {
    valores2 = [
        [

// OBSERVACIONES

    formulario.getRange("D8").getDisplayValue(),


    // Categorizacion

    formulario.getRange("B21").getDisplayValue(),
    formulario.getRange("D21").getDisplayValue(),
    formulario.getRange("F21").getDisplayValue(),

    // ENTREVISTA


    formulario.getRange("B23").getDisplayValue(),
    formulario.getRange("D23").getDisplayValue(),
    formulario.getRange("F23").getDisplayValue(),
    formulario.getRange("H23").getDisplayValue(),

    formulario.getRange("B24").getDisplayValue(),
    formulario.getRange("D24").getDisplayValue(),
    formulario.getRange("F24").getDisplayValue(),
    formulario.getRange("H24").getDisplayValue(),

    formulario.getRange("B25").getDisplayValue(),
    formulario.getRange("D25").getDisplayValue(),
    formulario.getRange("F25").getDisplayValue(),
    formulario.getRange("H25").getDisplayValue(),

    formulario.getRange("B26").getDisplayValue(),
    formulario.getRange("D26").getDisplayValue(),
    formulario.getRange("F26").getDisplayValue(),
    formulario.getRange("H26").getDisplayValue(),

    formulario.getRange("B27").getDisplayValue(),
    formulario.getRange("D27").getDisplayValue(),
    formulario.getRange("F27").getDisplayValue(),
    formulario.getRange("H27").getDisplayValue(),

    formulario.getRange("B28").getDisplayValue(),
    formulario.getRange("D28").getDisplayValue(),
    formulario.getRange("F28").getDisplayValue(),
    formulario.getRange("H28").getDisplayValue(),

    // Accion CETEC

    formulario.getRange("J28").getDisplayValue(),


    // CONVIVIENTES/CONTACTOS

    formulario.getRange("B30").getDisplayValue(),
    formulario.getRange("D30").getDisplayValue(),

    
    formulario.getRange("B32").getDisplayValue(),
    formulario.getRange("C32").getDisplayValue(),
    formulario.getRange("D32").getDisplayValue(),
    formulario.getRange("E32").getDisplayValue(),
    formulario.getRange("F32").getDisplayValue(),
    formulario.getRange("G32").getDisplayValue(),
    formulario.getRange("H32").getDisplayValue(),
    formulario.getRange("I32").getDisplayValue(),

    formulario.getRange("B33").getDisplayValue(),
    formulario.getRange("C33").getDisplayValue(),
    formulario.getRange("D33").getDisplayValue(),
    formulario.getRange("E33").getDisplayValue(),
    formulario.getRange("F33").getDisplayValue(),
    formulario.getRange("G33").getDisplayValue(),
    formulario.getRange("H33").getDisplayValue(),
    formulario.getRange("I33").getDisplayValue(),

    formulario.getRange("B34").getDisplayValue(),
    formulario.getRange("C34").getDisplayValue(),
    formulario.getRange("D34").getDisplayValue(),
    formulario.getRange("E34").getDisplayValue(),
    formulario.getRange("F34").getDisplayValue(),
    formulario.getRange("G34").getDisplayValue(),
    formulario.getRange("H34").getDisplayValue(),
    formulario.getRange("I34").getDisplayValue(),

    formulario.getRange("B35").getDisplayValue(),
    formulario.getRange("C35").getDisplayValue(),
    formulario.getRange("D35").getDisplayValue(),
    formulario.getRange("E35").getDisplayValue(),
    formulario.getRange("F35").getDisplayValue(),
    formulario.getRange("G35").getDisplayValue(),
    formulario.getRange("H35").getDisplayValue(),
    formulario.getRange("I35").getDisplayValue(),

    formulario.getRange("B36").getDisplayValue(),
    formulario.getRange("C36").getDisplayValue(),
    formulario.getRange("D36").getDisplayValue(),
    formulario.getRange("E36").getDisplayValue(),
    formulario.getRange("F36").getDisplayValue(),
    formulario.getRange("G36").getDisplayValue(),
    formulario.getRange("H36").getDisplayValue(),
    formulario.getRange("I36").getDisplayValue(),

    formulario.getRange("B37").getDisplayValue(),
    formulario.getRange("C37").getDisplayValue(),
    formulario.getRange("D37").getDisplayValue(),
    formulario.getRange("E37").getDisplayValue(),
    formulario.getRange("F37").getDisplayValue(),
    formulario.getRange("G37").getDisplayValue(),
    formulario.getRange("H37").getDisplayValue(),
    formulario.getRange("I37").getDisplayValue(),

    formulario.getRange("B38").getDisplayValue(),
    formulario.getRange("C38").getDisplayValue(),
    formulario.getRange("D38").getDisplayValue(),
    formulario.getRange("E38").getDisplayValue(),
    formulario.getRange("F38").getDisplayValue(),
    formulario.getRange("G38").getDisplayValue(),
    formulario.getRange("H38").getDisplayValue(),
    formulario.getRange("I38").getDisplayValue(),

    formulario.getRange("B39").getDisplayValue(),
    formulario.getRange("C39").getDisplayValue(),
    formulario.getRange("D39").getDisplayValue(),
    formulario.getRange("E39").getDisplayValue(),
    formulario.getRange("F39").getDisplayValue(),
    formulario.getRange("G39").getDisplayValue(),
    formulario.getRange("H39").getDisplayValue(),
    formulario.getRange("I39").getDisplayValue(),

    formulario.getRange("B40").getDisplayValue(),
    formulario.getRange("C40").getDisplayValue(),
    formulario.getRange("D40").getDisplayValue(),
    formulario.getRange("E40").getDisplayValue(),
    formulario.getRange("F40").getDisplayValue(),
    formulario.getRange("G40").getDisplayValue(),
    formulario.getRange("H40").getDisplayValue(),
    formulario.getRange("I40").getDisplayValue(),

    formulario.getRange("B41").getDisplayValue(),
    formulario.getRange("C41").getDisplayValue(),
    formulario.getRange("D41").getDisplayValue(),
    formulario.getRange("E41").getDisplayValue(),
    formulario.getRange("F41").getDisplayValue(),
    formulario.getRange("G41").getDisplayValue(),
    formulario.getRange("H41").getDisplayValue(),
    formulario.getRange("I41").getDisplayValue(),

    // Accion CETEC

    formulario.getRange("J41").getDisplayValue(),


    // VACUNACIÓN / SUBSIDIO / SALUD MENTAL


    formulario.getRange("B43").getDisplayValue(),
    formulario.getRange("D43").getDisplayValue(),
    formulario.getRange("F43").getDisplayValue(),

    formulario.getRange("B44").getDisplayValue(),
    formulario.getRange("D44").getDisplayValue(),
    formulario.getRange("F44").getDisplayValue(),
    formulario.getRange("H44").getDisplayValue(),
    formulario.getRange("B45").getDisplayValue(),

    // Accion CETEC

    formulario.getRange("J45").getDisplayValue(),

    // GESTACIÓN / PUERPERIO

    formulario.getRange("B47").getDisplayValue(),
    formulario.getRange("D47").getDisplayValue(),
    formulario.getRange("F47").getDisplayValue(),

    // Accion CETEC

    formulario.getRange("J47").getDisplayValue(),

    // SEGUIMIENTO 1

    formulario.getRange("B49").getDisplayValue(),
    formulario.getRange("D49").getDisplayValue(),
    formulario.getRange("F49").getDisplayValue(),
    formulario.getRange("H49").getDisplayValue(),
    formulario.getRange("J49").getDisplayValue(),

    formulario.getRange("B50").getDisplayValue(),
    formulario.getRange("D50").getDisplayValue(),
    formulario.getRange("F50").getDisplayValue(),
    formulario.getRange("H50").getDisplayValue(),
    formulario.getRange("J50").getDisplayValue(),

    formulario.getRange("B51").getDisplayValue(),
    formulario.getRange("D51").getDisplayValue(),

        // SEGUIMIENTO 2

    formulario.getRange("B53").getDisplayValue(),
    formulario.getRange("D53").getDisplayValue(),
    formulario.getRange("F53").getDisplayValue(),
    formulario.getRange("H53").getDisplayValue(),
    formulario.getRange("J53").getDisplayValue(),

    formulario.getRange("B54").getDisplayValue(),
    formulario.getRange("D54").getDisplayValue(),
    formulario.getRange("F54").getDisplayValue(),
    formulario.getRange("H54").getDisplayValue(),
    formulario.getRange("J54").getDisplayValue(),

    formulario.getRange("B55").getDisplayValue(),
    formulario.getRange("D55").getDisplayValue(),

        // SEGUIMIENTO 3

    formulario.getRange("B57").getDisplayValue(),
    formulario.getRange("D57").getDisplayValue(),
    formulario.getRange("F57").getDisplayValue(),
    formulario.getRange("H57").getDisplayValue(),
    formulario.getRange("J57").getDisplayValue(),

    formulario.getRange("B58").getDisplayValue(),
    formulario.getRange("D58").getDisplayValue(),
    formulario.getRange("F58").getDisplayValue(),
    formulario.getRange("H58").getDisplayValue(),
    formulario.getRange("J58").getDisplayValue(),

    formulario.getRange("B59").getDisplayValue(),
    formulario.getRange("D59").getDisplayValue(),

        // SEGUIMIENTO 4

    formulario.getRange("B61").getDisplayValue(),
    formulario.getRange("D61").getDisplayValue(),
    formulario.getRange("F61").getDisplayValue(),
    formulario.getRange("H61").getDisplayValue(),
    formulario.getRange("J61").getDisplayValue(),

    formulario.getRange("B62").getDisplayValue(),
    formulario.getRange("D62").getDisplayValue(),
    formulario.getRange("F62").getDisplayValue(),
    formulario.getRange("H62").getDisplayValue(),
    formulario.getRange("J62").getDisplayValue(),

    formulario.getRange("B63").getDisplayValue(),
    formulario.getRange("D63").getDisplayValue(),

        // SEGUIMIENTO 5

    formulario.getRange("B65").getDisplayValue(),
    formulario.getRange("D65").getDisplayValue(),
    formulario.getRange("F65").getDisplayValue(),
    formulario.getRange("H65").getDisplayValue(),
    formulario.getRange("J65").getDisplayValue(),

    formulario.getRange("B66").getDisplayValue(),
    formulario.getRange("D66").getDisplayValue(),
    formulario.getRange("F66").getDisplayValue(),
    formulario.getRange("H66").getDisplayValue(),
    formulario.getRange("J66").getDisplayValue(),

    formulario.getRange("B67").getDisplayValue(),
    formulario.getRange("D67").getDisplayValue(),
  

    // Accion CETEC

    formulario.getRange("J68").getDisplayValue(),

            ],
    ];
}

// Actualizar datos
function Actualizar() {
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
    var formulario = hojaActiva.getSheetByName("Ope 1"); // Nombre de hoja del formulario
    var datos = hojaActiva.getSheetByName("Planilla de casos"); // Nombre de hoja donde se almacenan datos

    var valor = formulario.getRange("B3").getDisplayValue();
    var dni = formulario.getRange("B9").getDisplayValue();
    var valores = hojaActiva
        .getSheetByName("Planilla de casos")
        .getDataRange()
        .getDisplayValues(); // Nombre de hoja donde se almacenan datos

    for (var i = 0; i < valores.length; i++) {
        var fila = valores[ i ];
        if (fila[ NUM_COLUMNA_BUSQUEDA ] == valor && valor == dni) {
            var INT_R = i + 1;

            // Llama a la función para asignar valores a valores1
            asignarValores1(formulario);
            // Utiliza la variable valores1 para actualizar los datos
            datos.getRange(INT_R, 1, 1, 14).setValues(valores1);
              // Llama a la función para asignar valores a valores2
            asignarValores2(formulario);
            // Utiliza la variable valores1 para actualizar los datos
            datos.getRange(INT_R, 15, 1, 186).setValues(valores2);
        }
    }
    Limpiar();
}

function botonActualizar() {
    // Actualizar los datos en la hoja de cálculo
    validarCampos();
    crearNuevasColumnasOActualizarDatos();
    guardarDerivaciones("Ope 1");
    Actualizar();
    // monitoreoOpesData();
    SpreadsheetApp.getUi().alert("Datos actualizados correctamente");
}

function botonBuscar() {
    LimpiarMenosDNI();
    pegarInformacionUltimaFecha();
    Buscar();
    var dni = SpreadsheetApp.getActiveSpreadsheet()
        .getActiveSheet()
        .getRange("B3")
        .getValue();
    if (dni === "") {
        return SpreadsheetApp.getUi().alert(
            "Los datos fueron borrados correctamente. Ingrese un DNI para cargar el caso"
        );
    } else {
        return SpreadsheetApp.getUi().alert("Caso cargado correctamente");
    }
}

function LimpiarCasillas() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formS = ss.getSheetByName("Ope 1");

    var rangesToClearCASILLAS = [ "M3:M60" ];
    for (var i = 0; i < rangesToClearCASILLAS.length; i++) {
        formS.getRange(rangesToClearCASILLAS[ i ]).clearContent();
    }
}
