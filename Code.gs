// ─────────────────────────────────────────────────────────────────────────────
// Arte/Visual — Encuesta de experiencia
// Google Apps Script  |  Pegar en: Extensions > Apps Script del Sheet
//
// INSTRUCCIONES DE DEPLOY (hacer una sola vez):
//  1. Abrí el Google Sheet → Extensions → Apps Script
//  2. Borrá el contenido existente y pegá este archivo completo
//  3. Guardá (Ctrl+S)
//  4. Deploy → New deployment → Web app
//       Execute as: Me (tu cuenta Google)
//       Who has access: Anyone
//  5. Copiá la URL que aparece (termina en /exec)
//  6. Pegala en index.html donde dice SCRIPT_URL = '...'
//  7. En index.html también poné la URL pública del formulario en FORM_URL
// ─────────────────────────────────────────────────────────────────────────────

// ── Configuración ─────────────────────────────────────────────────────────────
const CONFIG = {
  // Mail que recibe notificación por cada respuesta recibida
  notificationEmail: 'contacto@artevisual.com',

  // Nombre de la hoja donde se guardan las respuestas (se crea sola si no existe)
  responseSheet: 'Respuestas',

  // Hoja 1 con lista de clientes (columna A = Nombre, columna B = Email)
  clientSheet: 'Hoja 1',

  // URL pública del formulario (la que vas a compartir con clientes)
  // Cambiala por donde tengás hosteado el index.html
  formUrl: 'https://TU_DOMINIO/index.html',

  // Asunto del mail de invitación a clientes
  inviteSubject: 'Te invitamos a compartir tu experiencia con Arte/Visual',
};

// ── Recepción de respuestas del formulario (POST) ──────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    guardarRespuesta(data);
    enviarNotificacion(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Recibe respuestas del formulario vía GET (evita problemas de CORS/redirect con doPost)
function doGet(e) {
  if (e.parameter.payload) {
    try {
      const data = JSON.parse(e.parameter.payload);
      guardarRespuesta(data);
      enviarNotificacion(data);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, status: 'Arte/Visual encuesta activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Guardar respuesta en el Sheet ──────────────────────────────────────────────
function guardarRespuesta(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.responseSheet);

  // Crea la hoja y cabecera si no existe
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.responseSheet);
    sheet.appendRow([
      'Fecha y hora',
      'Satisfacción (1-5)',
      'Lo que más valora',
      'Procesos de trabajo',
      'Mejoras sugeridas',
      'IP / referencia',
    ]);
    // Estilo de cabecera
    const header = sheet.getRange(1, 1, 1, 6);
    header.setBackground('#de0a14');
    header.setFontColor('#ffffff');
    header.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const scaleLabels = {
    1: 'Muy insatisfecho',
    2: 'Insatisfecho',
    3: 'Neutral',
    4: 'Satisfecho',
    5: 'Muy satisfecho',
  };

  const satisfaction = data.q1
    ? data.q1 + ' / 5 — ' + (scaleLabels[data.q1] || '')
    : '(sin responder)';

  sheet.appendRow([
    new Date(),
    satisfaction,
    (data.q2 || '').trim() || '(sin responder)',
    data.q3 || '(sin responder)',
    (data.q4 || '').trim() || '(sin comentarios)',
    data.ref || '',
  ]);

  // Auto-ajusta el ancho de columnas
  sheet.autoResizeColumns(1, 6);
}

// ── Notificación interna por mail ──────────────────────────────────────────────
function enviarNotificacion(data) {
  const scaleLabels = {
    1: 'Muy insatisfecho', 2: 'Insatisfecho', 3: 'Neutral',
    4: 'Satisfecho', 5: 'Muy satisfecho',
  };

  const sat = data.q1 ? data.q1 + ' / 5 (' + (scaleLabels[data.q1] || '') + ')' : '(sin responder)';

  const cuerpo = `
Nueva respuesta recibida — Encuesta Arte/Visual
===============================================

📊 Satisfacción con el servicio: ${sat}

💬 Lo que más valora:
${(data.q2 || '').trim() || '(sin responder)'}

⚙️ Procesos de trabajo:
${data.q3 || '(sin responder)'}

💡 Mejoras sugeridas:
${(data.q4 || '').trim() || '(sin comentarios)'}

-----------------------------------------------
Ver todas las respuestas en el Sheet:
https://docs.google.com/spreadsheets/d/1inffJI1iSHKRY9AYk_ztPRs4WEsaNTEFW0KzFrRh52Q/edit
  `.trim();

  MailApp.sendEmail({
    to: CONFIG.notificationEmail,
    subject: '📋 Nueva respuesta recibida — Encuesta Arte/Visual',
    body: cuerpo,
  });
}

// ── Envío masivo de invitaciones a clientes (ejecutar UNA SOLA VEZ) ────────────
//
// Cómo usar:
//  1. Asegurate de que la "Hoja 1" tenga:
//       Columna A = Nombre del cliente
//       Columna B = Email del cliente
//     (la fila 1 puede ser cabecera, se saltea automáticamente si no parece un email)
//  2. En el editor de Apps Script, seleccioná esta función en el dropdown
//     y presioná ▶ Run
//  3. Revisá el Log (View > Logs) para ver el estado de cada envío
//
function enviarEncuestaAClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.clientSheet);

  if (!sheet) {
    Logger.log('❌ No se encontró la hoja "' + CONFIG.clientSheet + '"');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let enviados = 0;
  let errores = 0;

  data.forEach((row, index) => {
    const nombre = (row[0] || '').toString().trim();
    const email  = (row[1] || '').toString().trim().toLowerCase();

    // Saltear filas sin email válido o que parezcan cabecera
    if (!email || !email.includes('@') || email === 'email') {
      Logger.log('⏭ Fila ' + (index + 1) + ' salteada: "' + nombre + '" — ' + email);
      return;
    }

    try {
      const cuerpoHtml = generarMailInvitacion(nombre);
      GmailApp.sendEmail(email, CONFIG.inviteSubject, '', {
        htmlBody: cuerpoHtml,
        name: 'Arte/Visual',
        replyTo: CONFIG.notificationEmail,
      });
      Logger.log('✅ Enviado a ' + nombre + ' <' + email + '>');
      enviados++;

      // Pequeña pausa para no superar cuotas de Gmail
      Utilities.sleep(300);
    } catch (err) {
      Logger.log('❌ Error enviando a ' + email + ': ' + err.message);
      errores++;
    }
  });

  Logger.log('\n──────────────────────────────────');
  Logger.log('✅ Enviados: ' + enviados);
  Logger.log('❌ Errores:  ' + errores);
  Logger.log('──────────────────────────────────');

  // Notificación de fin de proceso
  MailApp.sendEmail({
    to: CONFIG.notificationEmail,
    subject: '📤 Envío masivo completado — Arte/Visual',
    body: 'Se enviaron ' + enviados + ' invitaciones a la encuesta.\nErrores: ' + errores + '.',
  });
}

// ── Template HTML del mail de invitación ───────────────────────────────────────
function generarMailInvitacion(nombre) {
  const nombreMostrar = nombre || 'cliente';
  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
</head>
<body style="margin:0;padding:0;background:#f2ecea;font-family:'Helvetica Neue',Arial,sans-serif;color:#181311;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f2ecea;padding:40px 16px;">
    <tr>
      <td align="center">
        <table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;">

          <!-- Header -->
          <tr>
            <td style="background:#ffffff;border-radius:18px 18px 0 0;padding:32px 40px 24px;border-bottom:3px solid #de0a14;text-align:center;">
              <div style="font-size:24px;font-weight:900;letter-spacing:-0.02em;color:#181311;">
                Arte<span style="color:#de0a14;">/</span>Visual
              </div>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="background:#ffffff;padding:36px 40px;">
              <p style="margin:0 0 8px;font-size:13px;font-weight:700;letter-spacing:0.22em;text-transform:uppercase;color:#de0a14;">
                Encuesta de experiencia
              </p>
              <h1 style="margin:0 0 20px;font-size:28px;font-weight:800;line-height:1.15;color:#181311;">
                Hola, ${nombreMostrar}.<br/>Tu mirada importa.
              </h1>
              <p style="margin:0 0 16px;font-size:16px;line-height:1.65;color:#6a5f5d;">
                Nos encantaría saber cómo fue trabajar con nosotros. Tu feedback es el
                motor que nos impulsa a mejorar y a seguir ofreciéndote el mejor servicio.
              </p>
              <p style="margin:0 0 32px;font-size:16px;line-height:1.65;color:#6a5f5d;">
                Son solo <strong style="color:#181311;">4 preguntas</strong> —
                menos de 2 minutos. Tus respuestas son completamente confidenciales.
              </p>

              <!-- CTA button -->
              <table cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td align="center">
                    <a href="${CONFIG.formUrl}"
                       style="display:inline-block;background:#de0a14;color:#ffffff;
                              text-decoration:none;font-size:16px;font-weight:700;
                              letter-spacing:0.01em;padding:16px 40px;border-radius:100px;
                              box-shadow:0 8px 24px -8px rgba(222,10,20,0.5);">
                      Completar encuesta →
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#f8f4f3;border-radius:0 0 18px 18px;padding:20px 40px;text-align:center;">
              <p style="margin:0;font-size:12px;color:#a0938f;line-height:1.6;">
                © 2026 Arte<span style="color:#de0a14;">/</span>Visual · Todos los derechos reservados<br/>
                Si no querés recibir más comunicaciones, respondé este mail indicándolo.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
  `.trim();
}
