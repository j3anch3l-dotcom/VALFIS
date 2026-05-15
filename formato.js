/**
 * Genera el HTML del formato v8 basado en la imagen proporcionada.
 * @param {Object} r Objeto con los datos del trabajador y crédito.
 * @param {string|boolean} tipoPeriodo 'quincenal' o true para quincenal, cualquier otro para mensual.
 */
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

  var idImagenDrive = "TU_ID_DE_ARCHIVO_AQUI";
  var urlImagen = "https://drive.google.com/uc?export=view&id=" + idImagenDrive;

  var tipo = String(r.Tipo || "").trim().toLowerCase();
  var chkN = (tipo === "nuevo") ? "X" : "&nbsp;";
  var chkR = (tipo === "refinanciamiento" || tipo === "refinanciado") ? "X" : "&nbsp;";

  // Estilos base
  var F    = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;';
  var FB   = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;font-weight:bold;';
  var UL   = 'border-bottom:1px solid #000;';
  var SEC  = 'font-family:Arial,sans-serif;font-size:8.5pt;color:#000;font-weight:bold;text-align:center;padding:10px 0 4px 0;';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>'+
    'body{font-family:Arial,sans-serif;margin:8mm 12mm;line-height:1.15;color:#000;}'+
    'table{border-collapse:collapse;width:100%;margin-bottom:1px;}'+
    '.label{'+FB+'white-space:nowrap;padding-right:4px;vertical-align:bottom;}'+
    '.value{'+F+UL+'padding-left:4px;vertical-align:bottom;}'+
    '.chk-box{border:1px solid #000;width:30px;height:18px;text-align:center;display:inline-block;line-height:18px;vertical-align:middle;margin-left:8px;}'+
    '.header-text{'+FB+'font-size:10.5pt;margin:0;text-align:center;}'+
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
      '<td style="width:72%;vertical-align:top;">'+
        '<table>'+
          '<tr><td class="label" style="width:165px;">NOMBRE DE LA EMPRESA:</td><td class="value">ACÉRCATE A TU NÓMINA S.A.P.I. DE C.V.</td></tr>'+
          '<tr><td class="label">NOMBRE DEL VENDEDOR:</td><td class="value">'+(r.Nombre_Promotor||'')+'</td></tr>'+
        '</table>'+
        '<table style="margin-top:4px;">'+
          '<tr>'+
            '<td class="label" style="width:50px;">FECHA:</td>'+
            '<td style="'+UL+'text-align:center;width:35px;">'+fDia+'</td><td style="text-align:center;width:15px;vertical-align:bottom;">/</td>'+
            '<td style="'+UL+'text-align:center;width:35px;">'+fMes+'</td><td style="text-align:center;width:15px;vertical-align:bottom;">/</td>'+
            '<td style="'+UL+'text-align:center;width:55px;">'+fAnio+'</td>'+
            '<td style="width:30px;"></td>'+
          '</tr>'+
        '</table>'+
        '<table style="margin-top:8px;">'+
          '<tr>'+
            '<td class="label" style="width:60px;">NUEVO</td><td style="width:40px;"><span class="chk-box">'+chkN+'</span></td>'+
            '<td style="width:40px;"></td>'+
            '<td class="label" style="width:140px;">REFINANCIAMIENTO</td><td style="width:40px;"><span class="chk-box">'+chkR+'</span></td>'+
          '</tr>'+
        '</table>'+
      '</td>'+
      '<td style="width:28%;text-align:right;vertical-align:top;">'+
         '<div style="border:1px solid #000;width:140px;height:90px;display:inline-block;text-align:center;vertical-align:middle;">'+
            '<img src="'+urlImagen+'" style="max-width:125px;max-height:80px;margin-top:5px;">'+
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
      '<td class="label" style="width:135px;">TELÉFONO DE CASA:</td><td class="value" style="width:30%;">'+(r.Tel_Casa||'')+'</td>'+
      '<td class="label" style="padding-left:15px;width:135px;">TELÉFONO CELULAR:</td><td class="value">'+(r.Tel_Celular||'')+'</td>'+
    '</tr>'+
    '<tr><td class="label">SECRETARÍA EN LA QUE LABORA:</td><td class="value" colspan="3">'+(r.Secretaria||'')+'</td></tr>'+
    '<tr><td class="label">CARGO O PUESTO QUE OCUPA:</td><td class="value" colspan="3">'+(r.Cargo||'')+'</td></tr>'+
    '<tr>'+
      '<td class="label" style="width:135px;">CLAVE DE EMPLEADO:</td><td class="value">'+(r.Clave||'')+'</td>'+
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
  '<div style="border:1px solid #000; padding:12px; margin-top:15px;">'+
    '<div style="'+FB+'text-align:center; margin-bottom:15px; font-size:8.5pt;">'+
      'SELLO AUTORIZADO DEL CENTRO DE TRABAJO<br>'+
      'QUE CERTIFICA LOS DATOS DEL TRABAJADOR'+
    '</div>'+

    '<table style="margin-bottom:20px;">'+
      '<tr>'+
        '<td style="width:55%; text-align:center; vertical-align:bottom; padding-bottom:5px;">'+
          '<div style="width:85%; border-bottom:1px solid #000; margin: 0 auto; height:70px;"></div>'+
          '<div style="'+FB+'font-size:7.5pt; margin-top:5px;">FIRMA</div>'+
        '</td>'+
        '<td style="width:45%; text-align:right;">'+
          '<div style="border:1px solid #000; width:220px; height:125px; display:inline-block; position:relative;">'+
             '<span style="position:absolute; top:50%; left:50%; transform:translate(-50%, -50%); color:#DDD; font-size:11pt; font-weight:bold; width:100%; text-align:center;">SELLO DE CERTIFICADO</span>'+
          '</div>'+
        '</td>'+
      '</tr>'+
    '</table>'+

    '<table>'+
      '<tr><td class="label" style="width:180px;">NOMBRE DE QUIEN CERTIFICA:</td><td class="value">&nbsp;</td></tr>'+
      '<tr><td class="label">PUESTO:</td><td class="value">&nbsp;</td></tr>'+
      '<tr><td class="label">FECHA DE CERTIFICACIÓN:</td><td class="value">&nbsp;</td></tr>'+
    '</table>'+
  '</div>'+

  '</body></html>';
}

if (typeof module !== 'undefined') {
  module.exports = { generarHTMLFormato };
}
