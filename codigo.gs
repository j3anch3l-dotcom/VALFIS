// ============================================================
// CONFIGURACIÓN GLOBAL
// ============================================================
const SPREADSHEET_ID = '1dfqA-hUxsyqa8uEsCLSww7cQLxgUu8vQAN517i2IA7c';
const HOJA_QUINCENALES = 'Validaciones_Quincenales';
const HOJA_MENSUALES   = 'Validaciones_Mensuales';
const HOJA_PROMOTORES  = 'Promotores';
const HOJA_USUARIOS    = 'Usuarios'; // ✅ CORREGIDO: era 'Usuario'

const ENCABEZADOS = [
  'ID', 'Fecha_Registro', 'Tipo_Periodo', 'Nombre_Promotor', 'Num_Promotor',
  'Fecha', 'Tipo', 'Nombre_Trabajador', 'Domicilio', 'Tel_Casa',
  'Tel_Celular', 'Secretaria', 'Cargo', 'Clave', 'RFC',
  'Monto_Solicitado', 'Plazo', 'Descuento', 'Total_Pagar'
];

// ============================================================
// DIAGNÓSTICO — Ejecuta manualmente desde el editor para ver
// los nombres exactos de todas las hojas en el Log
// ============================================================
function diagnosticarHojas() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojas = ss.getSheets();
  hojas.forEach(function(h) {
    var nombre = h.getName();
    Logger.log('Hoja: "' + nombre + '" | Longitud: ' + nombre.length);
  });
}

// ============================================================
// PUNTO DE ENTRADA WEB APP
// ✅ CORREGIDO: usar createHtmlOutputFromFile (no Template)
//    porque los archivos HTML no usan scriptlets <? ?>
// ============================================================
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'form';
  // La URL base del exec se inyecta en el HTML via scriptlet <?= urlBase ?>
  var urlBase = ScriptApp.getService().getUrl();

  if (page === 'admin') {
    var tpl = HtmlService.createTemplateFromFile('Admin');
    tpl.urlBase = urlBase;
    return tpl.evaluate()
      .setTitle('Panel Administrativo – Fiscalía')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var tpl = HtmlService.createTemplateFromFile('Index');
  tpl.urlBase = urlBase;
  return tpl.evaluate()
    .setTitle('Validación Fiscalía – Captura de Crédito')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// OBTENER NOMBRE DEL PROMOTOR
// ============================================================
function getPromotor(numPromotor) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hoja = ss.getSheetByName(HOJA_PROMOTORES);
    if (!hoja) return { ok: false, msg: 'Hoja Promotores no encontrada' };

    var datos = hoja.getDataRange().getValues();
    var numBuscado = numPromotor.toString().trim();

    for (var i = 1; i < datos.length; i++) {
      var celdaID     = datos[i][1];
      var celdaNombre = datos[i][2];

      if (
        String(celdaID).trim() === numBuscado ||
        Number(celdaID) === Number(numBuscado)
      ) {
        if (celdaNombre) return { ok: true, nombre: String(celdaNombre).trim() };
      }
    }

    return { ok: false, msg: 'Promotor "' + numBuscado + '" no encontrado' };

  } catch (err) {
    return { ok: false, msg: 'Error: ' + err.message };
  }
}

// ============================================================
// GUARDAR REGISTRO
// ============================================================
function guardarRegistro(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var nombreHoja = data.tipoPeriodo === 'quincenal' ? HOJA_QUINCENALES : HOJA_MENSUALES;
    var hoja = ss.getSheetByName(nombreHoja);

    // Crear hoja con encabezados correctos si no existe
    if (!hoja) {
      hoja = ss.insertSheet(nombreHoja);
      hoja.appendRow(ENCABEZADOS);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
    } else {
      // Verificar que la fila 1 tenga encabezados; si no, insertarla
      var lastRow = hoja.getLastRow();
      if (lastRow === 0) {
        hoja.appendRow(ENCABEZADOS);
        hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
      } else {
        var primeraFila = hoja.getRange(1, 1, 1, 1).getValue();
        if (primeraFila !== 'ID') {
          hoja.insertRowBefore(1);
          hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);
          hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
        }
      }
    }

    // Obtener nombre del promotor DESDE EL SERVIDOR para garantizar que sea correcto
    var nombrePromotor = '';
    var resPromotor = getPromotor(String(data.numPromotor));
    if (resPromotor.ok) {
      nombrePromotor = resPromotor.nombre;
    } else {
      nombrePromotor = data.nombrePromotor || '';
    }

    // Fecha de registro generada en el servidor (no puede venir del cliente)
    var now = new Date();
    var dia     = String(now.getDate()).padStart(2, '0');
    var mes     = String(now.getMonth() + 1).padStart(2, '0');
    var anio    = now.getFullYear();
    var horas   = String(now.getHours()).padStart(2, '0');
    var minutos = String(now.getMinutes()).padStart(2, '0');
    var fechaRegistroStr = dia + '/' + mes + '/' + anio + ' ' + horas + ':' + minutos;

    var id = Utilities.getUuid().substring(0, 8).toUpperCase();

    // Orden exacto según ENCABEZADOS:
    // ID, Fecha_Registro, Tipo_Periodo, Nombre_Promotor, Num_Promotor,
    // Fecha, Tipo, Nombre_Trabajador, Domicilio, Tel_Casa,
    // Tel_Celular, Secretaria, Cargo, Clave, RFC,
    // Monto_Solicitado, Plazo, Descuento, Total_Pagar
    var fila = [
      id,
      fechaRegistroStr,
      data.tipoPeriodo,
      nombrePromotor,
      data.numPromotor,
      data.fecha,
      data.tipo,
      data.nombreTrabajador,
      data.domicilio,
      data.telCasa,
      data.telCelular,
      data.secretaria,
      data.cargo,
      data.clave,
      data.rfc,
      parseFloat(data.montoSolicitado) || 0,
      data.plazo,
      parseFloat(data.descuento) || 0,
      parseFloat(data.totalPagar) || 0
    ];

    hoja.appendRow(fila);
    return { ok: true, id: id };

  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// ============================================================
// AUTENTICACIÓN
// ✅ CORREGIDO: usa la constante HOJA_USUARIOS directamente
// ============================================================
function login(usuarioParam, contrasenaParam) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hojaUsuarios = ss.getSheetByName(HOJA_USUARIOS);

    if (!hojaUsuarios) {
      var todasLasHojas = ss.getSheets();
      var nombresHojas = todasLasHojas.map(function(h) { return '"' + h.getName() + '"'; }).join(', ');
      return { ok: false, msg: 'Hoja de usuarios no encontrada. Hojas disponibles: ' + nombresHojas };
    }

    var datos = hojaUsuarios.getDataRange().getValues();
    var userIngresado = String(usuarioParam).trim();
    var passIngresado = String(contrasenaParam).trim();

    for (var i = 1; i < datos.length; i++) {
      var userEnHoja = String(datos[i][0]).trim();
      var passEnHoja = String(datos[i][1]).trim();

      if (userEnHoja === userIngresado && passEnHoja === passIngresado) {
        return { ok: true, usuario: userIngresado };
      }
    }

    return { ok: false, msg: 'Usuario o contraseña incorrectos' };

  } catch (err) {
    return { ok: false, msg: 'Error: ' + err.message };
  }
}

// ============================================================
// OBTENER REGISTROS (PANEL ADMIN)
// ============================================================
function getRegistros(tipoPeriodo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var nombreHoja = tipoPeriodo === 'quincenal' ? HOJA_QUINCENALES : HOJA_MENSUALES;
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() <= 1) return { ok: true, registros: [] };

    var datos = hoja.getDataRange().getValues();
    var encabezados = datos[0];
    var registros = [];

    for (var i = 1; i < datos.length; i++) {
      var obj = {};
      for (var j = 0; j < encabezados.length; j++) {
        var val = datos[i][j];
        // Convertir fechas a string legible para que no lleguen como objeto al HTML
        if (val instanceof Date) {
          var dia  = String(val.getDate()).padStart(2, '0');
          var mes  = String(val.getMonth() + 1).padStart(2, '0');
          var anio = val.getFullYear();
          val = dia + '/' + mes + '/' + anio;
        }
        obj[encabezados[j]] = val;
      }
      registros.push(obj);
    }

    return { ok: true, registros: registros };

  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// ============================================================
// GENERAR PDF — recibe (registroId, tipoPeriodo)
// Esta es la función que llama google.script.run desde el HTML
// ============================================================
function generarPDF(registroId, tipoPeriodo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var nombreHoja = tipoPeriodo === 'quincenal' ? HOJA_QUINCENALES : HOJA_MENSUALES;
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return { ok: false, msg: 'Hoja no encontrada: ' + nombreHoja };

    // ── Buscar el registro por ID ──
    var datos = hoja.getDataRange().getValues();
    var encabezados = datos[0];
    var registro = null;

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === String(registroId).trim()) {
        registro = {};
        for (var j = 0; j < encabezados.length; j++) {
          var v = datos[i][j];
          if (v instanceof Date) {
            var dd = String(v.getDate()).padStart(2, '0');
            var mm = String(v.getMonth() + 1).padStart(2, '0');
            var yy = v.getFullYear();
            v = dd + '/' + mm + '/' + yy;
          }
          registro[encabezados[j]] = v;
        }
        break;
      }
    }

    if (!registro) return { ok: false, msg: 'Registro no encontrado: ' + registroId };

    // ── Generar HTML y convertir a PDF vía Google Doc temporal ──
    var htmlContent = generarHTMLFormato(registro, tipoPeriodo);

    var htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, 'temp_validacion');
    var docFile = Drive.Files.insert(
      { title: 'temp_validacion_' + registroId, mimeType: MimeType.GOOGLE_DOCS },
      htmlBlob
    );
    var docId = docFile.id;

    var pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF);
    pdfBlob.setName('Validacion_' + tipoPeriodo + '_' + registroId + '.pdf');

    var pdfFile = DriveApp.createFile(pdfBlob);
    DriveApp.getFileById(docId).setTrashed(true); // eliminar doc temporal

    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var pdfId = pdfFile.getId();
    var url = 'https://drive.google.com/file/d/' + pdfId + '/view';

    return { ok: true, url: url, fileId: pdfId };

  } catch (err) {
    return { ok: false, msg: 'Error generarPDF: ' + err.message };
  }
}


// ============================================================
// GENERAR HTML DEL FORMATO v8 (ESTILO "DESCARGA")
// ============================================================
function generarHTMLFormato(r, tipoPeriodo) {
  var esQuincenal = (tipoPeriodo === true || tipoPeriodo === 'quincenal');
  var PT = esQuincenal ? "QUINCENAL" : "MENSUAL";

  var fechaHoy = "";
  try {
    fechaHoy = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy");
  } catch (e) {
    var d = new Date();
    var dd = String(d.getDate()).padStart(2, '0');
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    var yyyy = d.getFullYear();
    fechaHoy = dd + "/" + mm + "/" + yyyy;
  }

  var partes = String(fechaHoy).split('/');
  var fDia = partes[0]||''; var fMes = partes[1]||''; var fAnio = partes[2]||'';

  var idImagenDrive = "1Cg9UeL40L8G6I5G_XyD2P9rS_O0yN0yX"; // ID de imagen fiscalía
  var urlImagen = "https://drive.google.com/uc?export=view&id=" + idImagenDrive;

  var tipo = String(r.Tipo || "").trim().toLowerCase();
  var chkN = (tipo === "nuevo") ? "X" : "&nbsp;";
  var chkR = (tipo === "refinanciamiento" || tipo === "refinanciado") ? "X" : "&nbsp;";

  // Estilos base
  var F    = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;';
  var FB   = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;font-weight:bold;';
  var UL   = 'border-bottom:1px solid #000;';
  var SEC  = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;font-weight:bold;text-align:center;padding:12px 0 8px 0;';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>'+
    'body{font-family:Arial,sans-serif;margin:5mm 12mm 10mm 12mm;line-height:1.15;color:#000;}'+
    'table{border-collapse:collapse;width:100%;margin-bottom:1px;}'+
    '.label{'+FB+'white-space:nowrap;padding-right:5px;vertical-align:bottom;}'+
    '.value{'+F+UL+'padding-left:5px;vertical-align:bottom;}'+
    '.chk-box{border:1px solid #000;width:25px;height:18px;text-align:center;display:inline-block;line-height:18px;vertical-align:middle;margin-left:5px;}'+
    '.header-text{'+FB+'font-size:10.5pt;margin:0;text-align:center;}'+
    '.no-ul{border-bottom:none !important;}'+
  '</style></head><body>'+

  /* ENCABEZADO */
  '<div style="text-align:center; margin-bottom:15px;">'+
    '<p class="header-text">DIRECCIÓN GENERAL DE RECURSOS HUMANOS</p>'+
    '<p class="header-text">FISCALÍA GENERAL DEL ESTADO DE MORELOS</p>'+
    '<div style="height:12px;"></div>'+
    '<p class="header-text">VISTO BUENO PARA EL OTORGAMIENTO DE CRÉDITO</p>'+
  '</div>'+

  /* SECCIÓN 1: DATOS DE LA EMPRESA */
  '<div style="'+SEC+'">DATOS DE LA EMPRESA</div>'+
  '<table>'+
    '<tr>'+
      '<td style="width:75%;vertical-align:top;">'+
        '<table>'+
          '<tr><td class="label" style="width:165px;">NOMBRE DE LA EMPRESA:</td><td class="value">ACÉRCATE A TU NÓMINA S.A.P.I. DE C.V.</td></tr>'+
          '<tr><td class="label">NOMBRE DEL VENDEDOR:</td><td class="value">'+(r.Nombre_Promotor||'')+'</td></tr>'+
        '</table>'+
        '<table style="margin-top:8px;">'+
          '<tr>'+
            '<td class="label" style="width:50px;">FECHA:</td>'+
            '<td style="'+UL+'text-align:center;width:40px;">'+fDia+'</td><td style="text-align:center;width:20px;vertical-align:bottom;">/</td>'+
            '<td style="'+UL+'text-align:center;width:40px;">'+fMes+'</td><td style="text-align:center;width:20px;vertical-align:bottom;">/</td>'+
            '<td style="'+UL+'text-align:center;width:60px;">'+fAnio+'</td>'+
            '<td style="width:30px;"></td>'+
          '</tr>'+
        '</table>'+
        '<table style="margin-top:10px;">'+
          '<tr>'+
            '<td class="label" style="width:55px;">NUEVO</td><td style="width:40px;"><span class="chk-box">'+chkN+'</span></td>'+
            '<td style="width:40px;"></td>'+
            '<td class="label" style="width:130px;">REFINANCIAMIENTO</td><td style="width:40px;"><span class="chk-box">'+chkR+'</span></td>'+
          '</tr>'+
        '</table>'+
      '</td>'+
      '<td style="width:25%;text-align:right;vertical-align:top;">'+
         '<div style="border:1px solid #000;width:150px;height:100px;display:inline-block;text-align:center;vertical-align:middle;">'+
            '<img src="'+urlImagen+'" style="max-width:140px;max-height:90px;margin-top:5px;">'+
         '</div>'+
      '</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 2: DATOS DEL TRABAJADOR */
  '<div style="'+SEC+'">DATOS DEL TRABAJADOR</div>'+
  '<table>'+
    '<tr><td class="label" style="width:175px;">NOMBRE DEL TRABAJADOR:</td><td class="value" colspan="3">'+(r.Nombre_Trabajador||'')+'</td></tr>'+
    '<tr><td class="label">DOMICILIO PARTICULAR:</td><td class="value" colspan="3">'+(r.Domicilio||'')+'</td></tr>'+
    '<tr>'+
      '<td class="label" style="width:140px;">TELÉFONO DE CASA:</td><td class="value" style="width:30%;">'+(r.Tel_Casa||'')+'</td>'+
      '<td class="label" style="padding-left:15px;width:140px;">TELÉFONO CELULAR:</td><td class="value">'+(r.Tel_Celular||'')+'</td>'+
    '</tr>'+
    '<tr><td class="label">SECRETARÍA EN LA QUE LABORA:</td><td class="value" colspan="3">'+(r.Secretaria||'')+'</td></tr>'+
    '<tr><td class="label">CARGO O PUESTO QUE OCUPA:</td><td class="value" colspan="3">'+(r.Cargo||'')+'</td></tr>'+
    '<tr>'+
      '<td class="label" style="width:140px;">CLAVE DE EMPLEADO:</td><td class="value">'+(r.Clave||'')+'</td>'+
      '<td class="label" style="padding-left:15px;width:40px;">RFC:</td><td class="value">'+(r.RFC||'')+'</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 3: DATOS DEL CRÉDITO */
  '<div style="'+SEC+'">DATOS DEL CRÉDITO</div>'+
  '<table>'+
    '<tr>'+
      '<td class="label" style="width:150px;">MONTO SOLICITADO: $</td><td class="value" style="width:38%;">'+(r.Monto_Solicitado||'')+'</td>'+
      '<td class="label" style="padding-left:15px;width:60px;">PLAZO:</td><td class="value">'+(r.Plazo||'')+'</td>'+
    '</tr>'+
    '<tr>'+
      '<td class="label">DESCUENTO '+PT+': $</td><td class="value" style="width:38%;">'+(r.Descuento||'')+'</td>'+
      '<td colspan="2"></td>'+
    '</tr>'+
    '<tr>'+
      '<td class="label">TOTAL A PAGAR: $</td><td class="value" style="width:38%;">'+(r.Total_Pagar||'')+'</td>'+
      '<td class="label" style="padding-left:15px;width:135px;">INTERES MENSUAL:</td>'+
      '<td class="value">3.0 % global + IVA</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 4: PARA SER LLENADO POR LA DIRECCIÓN TÉCNICA DE PERSONAL */
  '<div style="'+SEC+'">PARA SER LLENADO POR LA DIRECCIÓN TÉCNICA DE PERSONAL</div>'+
  '<table>'+
    '<tr>'+
      '<td class="label" style="width:175px;">PERCEPCIÓN '+PT+': $</td><td class="value" style="width:30%;">&nbsp;</td>'+
      '<td class="label" style="padding-left:15px;width:175px;">DEDUCCIÓN '+PT+': $</td><td class="value">&nbsp;</td>'+
    '</tr>'+
    '<tr>'+
      '<td class="label">SUELDO NETO: $</td><td class="value">&nbsp;</td>'+
      '<td class="label" style="padding-left:15px;">FECHA DE INGRESO:</td><td class="value">&nbsp;</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 5: SELLOS Y FIRMAS */
  '<div style="border:1px solid #000; padding:15px; margin-top:15px;">'+
    '<div style="'+FB+'text-align:center; margin-bottom:20px; font-size:8.5pt;">'+
      'SELLO AUTORIZADO DEL CENTRO DE TRABAJO<br>'+
      'QUE CERTIFICA LOS DATOS DEL TRABAJADOR'+
    '</div>'+

    '<table style="margin-bottom:15px;">'+
      '<tr>'+
        '<td style="width:60%; text-align:center; vertical-align:bottom; padding-bottom:5px;">'+
          '<div style="width:85%; border-bottom:1px solid #000; margin: 0 auto; height:80px;"></div>'+
          '<div style="'+FB+'font-size:8.5pt; margin-top:8px;">FIRMA</div>'+
        '</td>'+
        '<td style="width:40%; text-align:right;">'+
          '<div style="border:1px solid #000; width:220px; height:135px; display:inline-block; position:relative; text-align:center; vertical-align:middle;">'+
             '<span style="position:absolute; top:50%; left:50%; transform:translate(-50%, -50%); color:#DDD; font-size:11pt; font-weight:bold; width:100%;">SELLO DE CERTIFICADO</span>'+
          '</div>'+
        '</td>'+
      '</tr>'+
    '</table>'+
  '</div>'+

  '<table style="margin-top:10px;">'+
    '<tr><td class="label" style="width:180px;">NOMBRE DE QUIEN CERTIFICA:</td><td class="value">&nbsp;</td></tr>'+
    '<tr><td class="label">PUESTO:</td><td class="value">&nbsp;</td></tr>'+
    '<tr><td class="label">FECHA DE CERTIFICACIÓN:</td><td class="value">&nbsp;</td></tr>'+
  '</table>'+

  '</body></html>';
}

// ============================================================
// REPARAR ENCABEZADOS DE HOJAS EXISTENTES
// Ejecuta esta función UNA VEZ manualmente desde el editor
// para corregir las hojas que ya tienen datos mal estructurados
// ============================================================
function repararEncabezados() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojas = [HOJA_MENSUALES, HOJA_QUINCENALES];

  hojas.forEach(function(nombreHoja) {
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      Logger.log('Hoja no encontrada: ' + nombreHoja);
      return;
    }

    var primeraFila = hoja.getRange(1, 1, 1, 1).getValue();
    Logger.log('Hoja: ' + nombreHoja + ' | Primera celda: "' + primeraFila + '"');

    if (primeraFila !== 'ID') {
      // Insertar fila de encabezados al inicio
      hoja.insertRowBefore(1);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
      Logger.log('✅ Encabezados insertados en: ' + nombreHoja);
    } else {
      Logger.log('✅ Encabezados ya correctos en: ' + nombreHoja);
    }
  });

  Logger.log('Reparación completada.');
}
