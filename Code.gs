// ============================================================
//  SMASH OR PASS — Google Apps Script Backend
// ============================================================
//  Estructura del Google Sheet (3 pestañas):
//
//  Pestaña "Users"
//    Fila 1 (encabezado): Name
//    Col A: nombre de cada jugadora
//
//  Pestaña "Celebs"
//    Fila 1 (encabezado): Name | Image URL
//    Col A: nombre del famoso/famosa
//    Col B: URL de imagen (opcional — si está vacío, la app
//            busca la foto en Wikipedia automáticamente)
//
//  Pestaña "Votes"
//    Fila 1 (encabezado): Round | Celebrity | Username | Vote | Timestamp
//    ← Esta pestaña se rellena automáticamente, no la toques.
//
// ─── CÓMO DESPLEGAR ─────────────────────────────────────────
//  1. Abre tu script en script.google.com
//  2. Pega todo este código (reemplaza lo que hubiera)
//  3. Haz click en "Implementar" → "Administrar implementaciones"
//  4. Edita la implementación existente (o crea una nueva):
//       Tipo:          Aplicación web
//       Ejecutar como: Yo (me)
//       Quién accede:  Cualquier persona
//  5. Implementar → autoriza permisos si te lo pide
//  6. Copia la URL y pégala en index.html (const API = '...')
// ============================================================

function doGet(e) {
  const p = e.parameter;
  try {
    switch (p.action) {
      case 'getUsers':        return ok(getUsers());
      case 'getCelebrities':  return ok(getCelebrities());
      case 'getVotes':        return ok(getVotes(p.round));
      case 'submitVote':      return ok(submitVote(p));
      default: return ok({ error: 'Acción desconocida: ' + p.action });
    }
  } catch (err) {
    return ok({ error: err.message });
  }
}

// Helper: respuesta JSON con CORS
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper: obtener una pestaña por nombre
function tab(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error(`Pestaña "${name}" no encontrada en el Sheet`);
  return s;
}

// ─── GET USERS ────────────────────────────────────────────────
function getUsers() {
  const rows  = tab('Users').getDataRange().getValues();
  const users = rows
    .slice(1)                              // saltar encabezado
    .map(r => (r[0] || '').toString().trim())
    .filter(Boolean);
  return { users };
}

// ─── GET CELEBRITIES ──────────────────────────────────────────
function getCelebrities() {
  const rows = tab('Celebs').getDataRange().getValues();
  const celebrities = rows
    .slice(1)
    .filter(r => r[0])
    .map(r => ({
      name:  (r[0] || '').toString().trim(),
      image: (r[1] || '').toString().trim()   // puede estar vacío
    }));
  return { celebrities };
}

// ─── GET VOTES FOR A ROUND ────────────────────────────────────
function getVotes(round) {
  if (!round) return { votes: [] };
  const rows  = tab('Votes').getDataRange().getValues();
  const votes = rows
    .slice(1)
    .filter(r => String(r[0]) === String(round))
    .map(r => ({
      round:     String(r[0]),
      celebrity: String(r[1] || ''),
      username:  String(r[2] || ''),
      vote:      String(r[3] || '')
    }));
  return { votes };
}

// ─── SUBMIT A VOTE ────────────────────────────────────────────
function submitVote(p) {
  const { round, celebrity, username, vote } = p;
  if (!round || !celebrity || !username || !vote) {
    return { error: 'Faltan parámetros' };
  }

  const s    = tab('Votes');
  const rows = s.getDataRange().getValues();

  // Evitar votos duplicados (misma ronda + celebridad + usuario)
  const dup = rows.slice(1).some(r =>
    String(r[0]) === String(round)                    &&
    String(r[1]).toLowerCase() === celebrity.toLowerCase() &&
    String(r[2]).toLowerCase() === username.toLowerCase()
  );

  if (!dup) {
    s.appendRow([round, celebrity, username, vote, new Date().toISOString()]);
  }

  return { success: true, alreadyVoted: dup };
}
