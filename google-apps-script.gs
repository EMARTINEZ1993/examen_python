const SHEET_NAME = "Resultados";

function doGet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService
      .createTextOutput("[]")
      .setMimeType(ContentService.MimeType.JSON);
  }

  const rows = sheet.getDataRange().getValues().slice(1);
  const attempts = rows.map(row => ({
    name: String(row[0] || ""),
    group: String(row[1] || ""),
    correct: Number(row[2] || 0),
    total: Number(row[3] || 15),
    score: String(row[4] || "0.0"),
    date: String(row[5] || "")
  }));

  return ContentService
    .createTextOutput(JSON.stringify(attempts))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService
      .createTextOutput("Esta función se ejecuta desde el examen. Para probarla manualmente, ejecuta testDoPost().")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const data = JSON.parse(e.postData.contents);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Nombre", "Grupo", "Aciertos", "Total", "Nota", "Fecha"]);
  }

  sheet.appendRow([
    data.name,
    data.group,
    data.correct,
    data.total,
    data.score,
    data.date
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function testDoPost() {
  doPost({
    postData: {
      contents: JSON.stringify({
        name: "Prueba docente",
        group: "G1",
        correct: 15,
        total: 15,
        score: "5.0",
        date: new Date().toLocaleString()
      })
    }
  });
}
