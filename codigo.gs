// ============================================================
// CONFIGURACIÓN GLOBAL
// ============================================================
const SPREADSHEET_ID = '1dfqA-hUxsyqa8uEsCLSww7cQLxgUu8vQAN517i2IA7c';
const HOJA_QUINCENALES = 'Validaciones_Quincenales';
const HOJA_MENSUALES   = 'Validaciones_Mensuales';
const HOJA_PROMOTORES  = 'Promotores';
const HOJA_USUARIOS    = 'Usuarios';

const ENCABEZADOS = [
  'ID', 'Fecha_Registro', 'Tipo_Periodo', 'Nombre_Promotor', 'Num_Promotor',
  'Fecha', 'Tipo', 'Nombre_Trabajador', 'Domicilio', 'Tel_Casa',
  'Tel_Celular', 'Secretaria', 'Cargo', 'Clave', 'RFC',
  'Monto_Solicitado', 'Plazo', 'Descuento', 'Total_Pagar'
];

function diagnosticarHojas() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojas = ss.getSheets();
  hojas.forEach(function(h) {
    var nombre = h.getName();
    Logger.log('Hoja: "' + nombre + '" | Longitud: ' + nombre.length);
  });
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'form';
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

function guardarRegistro(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var nombreHoja = data.tipoPeriodo === 'quincenal' ? HOJA_QUINCENALES : HOJA_MENSUALES;
    var hoja = ss.getSheetByName(nombreHoja);

    if (!hoja) {
      hoja = ss.insertSheet(nombreHoja);
      hoja.appendRow(ENCABEZADOS);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
    } else {
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

    var nombrePromotor = '';
    var resPromotor = getPromotor(String(data.numPromotor));
    if (resPromotor.ok) {
      nombrePromotor = resPromotor.nombre;
    } else {
      nombrePromotor = data.nombrePromotor || '';
    }

    var now = new Date();
    var dia     = String(now.getDate()).padStart(2, '0');
    var mes     = String(now.getMonth() + 1).padStart(2, '0');
    var anio    = now.getFullYear();
    var horas   = String(now.getHours()).padStart(2, '0');
    var minutos = String(now.getMinutes()).padStart(2, '0');
    var fechaRegistroStr = dia + '/' + mes + '/' + anio + ' ' + horas + ':' + minutos;

    var id = Utilities.getUuid().substring(0, 8).toUpperCase();

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

function generarPDF(registroId, tipoPeriodo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var nombreHoja = tipoPeriodo === 'quincenal' ? HOJA_QUINCENALES : HOJA_MENSUALES;
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return { ok: false, msg: 'Hoja no encontrada: ' + nombreHoja };

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
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    DriveApp.getFileById(docId).setTrashed(true);

    var pdfId = pdfFile.getId();
    var url = 'https://drive.google.com/file/d/' + pdfId + '/view';

    return { ok: true, url: url, fileId: pdfId };

  } catch (err) {
    return { ok: false, msg: 'Error generarPDF: ' + err.message };
  }
}


// ============================================================
// GENERAR HTML DEL FORMATO v23 (Ajuste visual final según image.png)
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

  var idImagenDrive = "1XPNH4HUSvms7IDwg8g1-95JLs_J1bU4b";
  var urlImagen = "https://drive.google.com/uc?export=view&id=" + idImagenDrive;

  var tipo = String(r.Tipo || "").trim().toLowerCase();
  var chkN = (tipo === "nuevo") ? "X" : "&nbsp;&nbsp;";
  var chkR = (tipo === "refinanciamiento" || tipo === "refinanciado") ? "X" : "&nbsp;&nbsp;";

  var F    = 'font-family:Arial,sans-serif;font-size:8.2pt;color:#000;';
  var FB   = 'font-family:Arial,sans-serif;font-size:8.2pt;color:#000;font-weight:bold;';
  var UL   = 'border-bottom:1px solid #000;';
  var SEC  = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;font-weight:bold;text-align:center;padding:10px 0 2px 0;';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>'+
    'body{font-family:Arial,sans-serif;margin:0;padding:5mm 15mm;line-height:1.05;color:#000;}'+
    'table{border-collapse:collapse;width:100%;margin-bottom:0px;border:none;}'+
    'td{padding:1.2px 0; border:none;}'+
    '.label{'+FB+'white-space:nowrap;padding-right:5px;vertical-align:bottom;}'+
    '.value{'+F+'border-bottom:1px solid #000;padding-left:5px;vertical-align:bottom;}'+
    '.chk-box{border:1px solid #000; width:12px; height:12px; display:inline-block; text-align:center; line-height:12px; font-weight:bold; font-size:8pt;}'+
  '</style></head><body>'+

  /* ENCABEZADO */
  '<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-bottom:5px;">'+
    '<tr><td align="center" style="'+FB+'font-size:10pt;">DIRECCIÓN GENERAL DE RECURSOS HUMANOS</td></tr>'+
    '<tr><td align="center" style="'+FB+'font-size:10pt;">FISCALÍA GENERAL DEL ESTADO DE MORELOS</td></tr>'+
    '<tr><td height="5"></td></tr>'+
    '<tr><td align="center" style="'+FB+'font-size:9pt;text-decoration:underline;">VISTO BUENO PARA EL OTORGAMIENTO DE CRÉDITO</td></tr>'+
    '<tr><td height="5"></td></tr>'+
  '</table>'+

  /* SECCIÓN 1: DATOS DE LA EMPRESA */
  '<div style="'+SEC+'">DATOS DE LA EMPRESA</div>'+
  '<table width="100%" border="0" cellspacing="0" cellpadding="0">'+
    '<tr>'+
      '<td width="65%" valign="top" style="padding-right:40px;">'+
        '<table width="100%" border="0" cellspacing="0" cellpadding="1">'+
          '<tr><td style="'+FB+'" width="150" height="22">NOMBRE DE LA EMPRESA:</td><td style="'+F+UL+'">ACÉRCATE A TU NÓMINA S.A.P.I. DE C.V.</td></tr>'+
          '<tr><td style="'+FB+'" height="22">NOMBRE DEL VENDEDOR:</td><td style="'+F+UL+'">'+(r.Nombre_Promotor||'')+'</td></tr>'+
        '</table>'+
        '<table width="100%" border="0" cellspacing="0" cellpadding="1" style="margin-top:3px;">'+
          '<tr>'+
            '<td style="'+FB+'" width="35">FECHA:</td>'+
            '<td style="'+F+UL+'" width="25" align="center">'+fDia+'</td><td width="10" align="center" valign="bottom">/</td>'+
            '<td style="'+F+UL+'" width="25" align="center">'+fMes+'</td><td width="10" align="center" valign="bottom">/</td>'+
            '<td style="'+F+UL+'" width="35" align="center">'+fAnio+'</td>'+
            '<td width="30"></td>'+
          '</tr>'+
        '</table>'+
        '<table width="100%" border="0" cellspacing="0" cellpadding="1" style="margin-top:5px;">'+
          '<tr>'+
            '<td style="'+FB+'" width="40">NUEVO</td><td width="25" align="center"><span class="chk-box">'+chkN+'</span></td>'+
            '<td width="45"></td>'+
            '<td style="'+FB+'" width="110">REFINANCIAMIENTO</td><td width="25" align="center"><span class="chk-box">'+chkR+'</span></td>'+
          '</tr>'+
        '</table>'+
      '</td>'+
      '<td width="35%" align="right" valign="top">'+
         '<table border="1" cellspacing="0" cellpadding="0" style="border:1px solid #000; width:140px; height:70px; border-collapse:collapse;">'+
            '<tr><td align="center" valign="middle" height="70" style="border:1px solid #000;">'+
               '<img src="'+urlImagen+'" width="130" height="60" style="display:block; margin:auto;">'+
            '</td></tr>'+
         '</table>'+
      '</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 2: DATOS DEL TRABAJADOR */
  '<div style="'+SEC+'">DATOS DEL TRABAJADOR</div>'+
  '<table width="100%" border="0" cellspacing="0" cellpadding="1">'+
    '<tr><td style="'+FB+'" width="160">NOMBRE DEL TRABAJADOR:</td><td style="'+F+UL+'" colspan="3">'+(r.Nombre_Trabajador||'')+'</td></tr>'+
    '<tr><td style="'+FB+'">DOMICILIO PARTICULAR:</td><td style="'+F+UL+'" colspan="3">'+(r.Domicilio||'')+'</td></tr>'+
    '<tr>'+
      '<td style="'+FB+'" width="120">TELÉFONO DE CASA:</td><td style="'+F+UL+'" width="30%">'+(r.Tel_Casa||'')+'</td>'+
      '<td style="'+FB+'padding-left:15px;" width="120">TELÉFONO CELULAR:</td><td style="'+F+UL+'">'+(r.Tel_Celular||'')+'</td>'+
    '</tr>'+
    '<tr><td style="'+FB+'">SECRETARÍA EN LA QUE LABORA:</td><td style="'+F+UL+'" colspan="3">'+(r.Secretaria||'')+'</td></tr>'+
    '<tr><td style="'+FB+'">CARGO O PUESTO QUE OCUPA:</td><td style="'+F+UL+'" colspan="3">'+(r.Cargo||'')+'</td></tr>'+
    '<tr>'+
      '<td style="'+FB+'" width="120">CLAVE DE EMPLEADO:</td><td style="'+F+UL+'">'+(r.Clave||'')+'</td>'+
      '<td style="'+FB+'padding-left:15px;" width="35">RFC:</td><td style="'+F+UL+'">'+(r.RFC||'')+'</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 3: DATOS DEL CRÉDITO */
  '<div style="'+SEC+'">DATOS DEL CRÉDITO</div>'+
  '<table width="100%" border="0" cellspacing="0" cellpadding="1">'+
    '<tr>'+
      '<td style="'+FB+'" width="130">MONTO SOLICITADO: $</td><td style="'+F+UL+'" width="38%">'+(r.Monto_Solicitado||'')+'</td>'+
      '<td style="'+FB+'padding-left:15px;" width="50">PLAZO:</td><td style="'+F+UL+'">'+(r.Plazo||'')+'</td>'+
    '</tr>'+
    '<tr>'+
      '<td style="'+FB+'">DESCUENTO '+PT+': $</td><td style="'+F+UL+'" width="38%">'+(r.Descuento||'')+'</td>'+
      '<td colspan="2"></td>'+
    '</tr>'+
    '<tr>'+
      '<td style="'+FB+'">TOTAL A PAGAR: $</td><td style="'+F+UL+'" width="38%">'+(r.Total_Pagar||'')+'</td>'+
      '<td style="'+FB+'padding-left:15px;" width="120">INTERES MENSUAL:</td>'+
      '<td style="'+F+UL+'">3.0 % global + IVA</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 4: PARA SER LLENADO POR LA DT */
  '<div style="'+SEC+'">PARA SER LLENADO POR LA DIRECCIÓN TÉCNICA DE PERSONAL</div>'+
  '<table width="100%" border="0" cellspacing="0" cellpadding="1">'+
    '<tr>'+
      '<td style="'+FB+'" width="155">PERCEPCIÓN '+PT+': $</td><td style="'+F+UL+'" width="30%">&nbsp;</td>'+
      '<td style="'+FB+'padding-left:15px;" width="155">DEDUCCIÓN '+PT+': $</td><td style="'+F+UL+'">&nbsp;</td>'+
    '</tr>'+
    '<tr>'+
      '<td style="'+FB+'">SUELDO NETO: $</td><td style="'+F+UL+'">&nbsp;</td>'+
      '<td style="'+FB+'padding-left:10px;">FECHA DE INGRESO:</td><td style="'+F+UL+'">&nbsp;</td>'+
    '</tr>'+
  '</table>'+

  /* SECCIÓN 5: SELLOS Y FIRMAS */
  '<div style="margin-top:10px;">'+
    '<table width="100%" border="0" cellspacing="0" cellpadding="0">'+
      '<tr><td align="center" style="'+FB+'">SELLO AUTORIZADO DEL CENTRO DE TRABAJO<br>QUE CERTIFICA LOS DATOS DEL TRABAJADOR</td></tr>'+
      '<tr><td height="20"></td></tr>'+
      '<tr>'+
        '<td>'+
          '<table width="100%" border="0" cellspacing="0" cellpadding="0">'+
            '<tr>'+
              '<td width="55%" align="center" valign="bottom" style="padding-bottom:10px;">'+
                '<div style="'+FB+'">FIRMA</div>'+
                '<div style="height:50px;"></div>'+
                '<div style="width:210px; border-bottom:1px solid #000; margin: 0 auto;"></div>'+
              '</td>'+
              '<td width="45%" align="right" valign="bottom">'+
                '<table border="1" cellspacing="0" cellpadding="0" style="border:1px solid #000; width:180px; height:110px; border-collapse:collapse;">'+
                  '<tr><td align="center" valign="middle" style="'+FB+' font-size:9.5pt; color:#000;">SELLO DE<br>CERTIFICADO</td></tr>'+
                '</table>'+
              '</td>'+
            '</tr>'+
          '</table>'+
        '</td>'+
      '</tr>'+
      '<tr><td height="20"></td></tr>'+
      '<tr>'+
        '<td>'+
          '<table width="100%" border="0" cellspacing="0" cellpadding="2">'+
            '<tr><td style="'+FB+'" width="165">NOMBRE DE QUIEN CERTIFICA:</td><td style="'+F+UL+'">&nbsp;</td></tr>'+
            '<tr><td style="'+FB+'">PUESTO:</td><td style="'+F+UL+'">&nbsp;</td></tr>'+
            '<tr><td style="'+FB+'">FECHA DE CERTIFICACIÓN:</td><td style="'+F+UL+'">&nbsp;</td></tr>'+
          '</table>'+
        '</td>'+
      '</tr>'+
    '</table>'+
  '</div>'+

  '</body></html>';
}

function repararEncabezados() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojas = [HOJA_MENSUALES, HOJA_QUINCENALES];

  hojas.forEach(function(nombreHoja) {
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;

    var primeraFila = hoja.getRange(1, 1, 1, 1).getValue();
    if (primeraFila !== 'ID') {
      hoja.insertRowBefore(1);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);
      hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
    }
  });
}
