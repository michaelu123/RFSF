// RFSF

interface MapS2I {
  [others: string]: number;
}
interface MapS2S {
  [others: string]: string;
}
interface HeaderMap {
  [others: string]: MapS2I;
}

let inited = false;
let headers: HeaderMap = {};
let kurseSheet: GoogleAppsScript.Spreadsheet.Sheet;
let buchungenSheet: GoogleAppsScript.Spreadsheet.Sheet;

// Indices are 1-based!!
// Buchungen
let mailIndex: number; // E-Mail-Adresse
let kursIndexB: number; // Welchen Kurs möchten Sie belegen?
let herrFrauIndex: number; // Anrede
let nameIndex: number; // Name
let mitgliedsNummerIndex: number; // ADFC-Mitgliedsnummer falls Mitglied
let zahlungsArtIndex: number; // Zahlungsart
let zustimmungsIndex: number; // Zustimmung zur SEPA-Lastschrift
let bestätigungsIndex: number; // Bestätigung (der Teilnahmebedingungen)
let verifikationsIndex: number; // Verifikation (der Email-Adresse)
let anmeldebestIndex: number; // Anmeldebestätigung (gesendet)
let bezahltIndex: number; // Bezahlt

// Kurse
let kursNameIndex: number; // Kursname
let tagIndex: number; // Datum
let ersatzIndex: number; // Datum
let uhrZeitIndex: number; // Uhrzeit
let kursOrtIndex: number; // Kursort
let anzahlIndex: number; // Kursplätze
let restIndex: number; // Restplätze

// map Buchungen headers to print headers
let printCols = new Map([
  ["Vorname", "Vorname"],
  ["Name", "Nachname"],
  ["Telefonnummer für Rückfragen", "Telefon"],
  ["Anrede", "Anrede"],
  ["E-Mail-Adresse", "Email"],
  ["Ort", "Ort"],
]);

const kursFrage = "Welchen Kurs möchten Sie belegen?";

interface Event {
  namedValues: { [others: string]: string[] };
  range: GoogleAppsScript.Spreadsheet.Range;
  [others: string]: any;
}

function isEmpty(str: string | undefined | null) {
  if (typeof str == "number") return false;
  return !str || 0 === str.length; // I think !str is sufficient...
}

function test() {
  init();
  let e: Event = {
    namedValues: {
      Vorname: ["Michael"],
      Name: ["Uhlenberg"],
      Anrede: ["Herr"],
      Zahlungsart: ["Überweisung"],
      "E-Mail-Adresse": ["michael.uhlenberg@t-online.de"],
      "Lastschrift: IBAN-Kontonummer": ["DE91100000000123456789"],
      [kursFrage]: ["RFS_1F01G"],
    },
    range: buchungenSheet.getRange(2, 1, 1, buchungenSheet.getLastColumn()),
  };
  dispatch(e);
}

function init() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  for (let sheet of sheets) {
    let sheetName = sheet.getName();
    let sheetHeaders: MapS2I = {};
    // Logger.log("sheetName %s", sheetName);
    headers[sheetName] = sheetHeaders;
    let numCols = sheet.getLastColumn();
    // Logger.log("numCols %s", numCols);
    let row1Vals = sheet.getRange(1, 1, 1, numCols).getValues();
    // Logger.log("sheetName %s row1 %s", sheetName, row1Vals);
    for (let i = 0; i < numCols; i++) {
      let v: string = row1Vals[0][i];
      if (isEmpty(v)) continue;
      sheetHeaders[v] = i + 1;
    }
    // Logger.log("sheet %s %s", sheetName, sheetHeaders);

    if (sheet.getName() == "Kurse") {
      kurseSheet = sheet;
      kursNameIndex = sheetHeaders["Kursname"];
      uhrZeitIndex = sheetHeaders["Uhrzeit"];
      tagIndex = sheetHeaders["Datum"];
      ersatzIndex = sheetHeaders["Ersatztermin"];
      kursOrtIndex = sheetHeaders["Kursort"];
      anzahlIndex = sheetHeaders["Kursplätze"];
      restIndex = sheetHeaders["Restplätze"];
    } else if (sheet.getName() == "Buchungen") {
      buchungenSheet = sheet;
      mailIndex = sheetHeaders["E-Mail-Adresse"];
      kursIndexB = sheetHeaders[kursFrage];
      herrFrauIndex = sheetHeaders["Anrede"];
      nameIndex = sheetHeaders["Name"];
      mitgliedsNummerIndex =
        sheetHeaders["ADFC-Mitgliedsnummer falls Mitglied"];
      nameIndex = sheetHeaders["Name"];
      zahlungsArtIndex = sheetHeaders["Zahlungsart"];
      zustimmungsIndex = sheetHeaders["Zustimmung zur SEPA-Lastschrift"];
      bestätigungsIndex = sheetHeaders["Bestätigung"];
      verifikationsIndex = sheetHeaders["Verifikation"];
      anmeldebestIndex = sheetHeaders["Anmeldebestätigung"];
      bezahltIndex = sheetHeaders["Bezahlt"];

      if (verifikationsIndex == null) {
        verifikationsIndex = addColumn(sheet, sheetHeaders, "Verifikation");
      }
      if (anmeldebestIndex == null) {
        anmeldebestIndex = addColumn(sheet, sheetHeaders, "Anmeldebestätigung");
      }
      if (bezahltIndex == null) {
        bezahltIndex = addColumn(sheet, sheetHeaders, "Bezahlt");
      }
    }
    inited = true;
  }
}

// add a cell in row 1 with a new column title, return its index
function addColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetHeaders: MapS2I,
  title: string,
): number {
  let max = 0;
  for (let sh in sheetHeaders) {
    if (sheetHeaders[sh] > max) max = sheetHeaders[sh];
  }
  if (max >= sheet.getMaxColumns()) {
    sheet.insertColumnAfter(max);
  }
  max += 1;
  sheet.getRange(1, max).setValue(title);
  sheetHeaders[title] = max;
  return max;
}

function anredeText(herrFrau: string, name: string) {
  if (herrFrau === "Herr") {
    return "Sehr geehrter Herr " + name;
  } else {
    return "Sehr geehrte Frau " + name;
  }
}

function heuteString() {
  return Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    "YYYY-MM-dd HH:mm:ss",
  );
}

function attachmentFiles() {
  let thisFileId = SpreadsheetApp.getActive().getId();
  let thisFile = DriveApp.getFileById(thisFileId);
  let parent = thisFile.getParents().next();
  let grandPa = parent.getParents().next();
  let attachmentFolder = grandPa
    .getFoldersByName("Texte für Fahrsicherheitstrainings")
    .next();
  let PDFs = attachmentFolder.getFilesByType("application/pdf"); // MimeType.PDF
  let files = [];
  while (PDFs.hasNext()) {
    files.push(PDFs.next());
  }
  return files; // why not use PDFs directly??
}

function kursPreis(kurs: string, mitgliedsNummer: string): number {
  if (kurs.endsWith("G")) return isEmpty(mitgliedsNummer) ? 30 : 15;
  if (kurs.endsWith("A")) return isEmpty(mitgliedsNummer) ? 40 : 20;
  if (kurs.endsWith("P")) return isEmpty(mitgliedsNummer) ? 20 : 10;
  return 9999;
}

function anmeldebestätigung() {
  if (!inited) init();
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Buchungen") {
    SpreadsheetApp.getUi().alert(
      "Bitte eine Zeile im Sheet 'Buchungen' selektieren",
    );
    return;
  }
  let curCell = sheet.getSelection().getCurrentCell();
  if (!curCell) {
    SpreadsheetApp.getUi().alert("Bitte zuerst Teilnehmerzeile selektieren");
    return;
  }
  let row = curCell.getRow();
  if (row < 2 || row > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Teilnehmerzeile selektieren",
    );
    return;
  }
  let rowValues = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  let rowNote = sheet.getRange(row, 1).getNote();
  if (!isEmpty(rowNote)) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile hat eine Notiz und ist deshalb ungültig",
    );
    return;
  }
  if (isEmpty(rowValues[verifikationsIndex - 1])) {
    SpreadsheetApp.getUi().alert("Email-Adresse nicht verifiziert");
    return;
  }
  if (!isEmpty(rowValues[anmeldebestIndex - 1])) {
    SpreadsheetApp.getUi().alert("Der Kurs wurde schon bestätigt");
    return;
  }
  // setting up mail
  let emailTo: string = rowValues[mailIndex - 1].toLowerCase().trim();
  let subject: string = "Bestätigung Ihrer Kursanmeldung";
  let herrFrau = rowValues[herrFrauIndex - 1];
  let name = rowValues[nameIndex - 1];
  // Anrede
  let anrede: string = anredeText(herrFrau, name);
  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    "emailBestätigung.html",
  );

  let kurs: string = rowValues[kursIndexB - 1];
  let kursDesc: string = "";
  if (kurs.endsWith("G")) kursDesc = "Grundkurs Fahrsicherheitstraining";
  if (kurs.endsWith("A")) kursDesc = "Aufbaukurs Fahrsicherheitstraining";
  if (kurs.endsWith("P")) kursDesc = "Pedelectraining für Senioren";
  let mitgliedsNummer: string = rowValues[mitgliedsNummerIndex - 1];

  let betrag: number = kursPreis(kurs, mitgliedsNummer);
  let einzug: boolean = rowValues[zahlungsArtIndex - 1].startsWith("SEPA");
  let zahlungsText: string;
  if (einzug) {
    zahlungsText =
      'Sie haben als Zahlungsart "SEPA Lastschrift" gewählt. Wir ziehen die Teilnahmegebühr von ' +
      betrag +
      "€ in den nächsten Tagen ein.";
  } else {
    zahlungsText =
      "Bitte überweisen Sie " +
      betrag +
      "€ auf das Konto DE62 7015 0000 0904 1577 81 bei der Stadtsparkasse München unter Angabe der Kursnummer.";
  }

  let kursRow = null;
  let kurseS: Array<Array<string>> = kurseSheet.getSheetValues(
    2,
    1,
    kurseSheet.getLastRow(),
    kurseSheet.getLastColumn(),
  );

  for (let j = 0; j < kurseS.length; j++) {
    if (kurseS[j][0] == kurs) {
      kursRow = kurseS[j];
      break;
    }
  }
  Logger.log("kursRow %s", kursRow);
  if (!kursRow) {
    SpreadsheetApp.getUi().alert("Kurs '" + kurs + "' nicht im Kurse-Sheet!?");
    return;
  }
  let ort: string = kursRow[kursOrtIndex - 1];
  let termin = any2Str(kursRow[tagIndex - 1], "E 'den' dd.MM", false);
  termin += " von " + kursRow[uhrZeitIndex - 1];
  Logger.log("termin %s", termin);

  template.anrede = anrede;
  template.kurs = kurs + " - " + kursDesc;
  template.ort = ort;
  template.termin = termin;
  template.zahlungstext = zahlungsText;

  SpreadsheetApp.getUi().alert(
    herrFrau + " " + name + " bucht den Kurs " + kurs,
  );

  let htmlText: string = template.evaluate().getContent();
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "Radfahrschule ADFC München e.V.",
    replyTo: "radfahrschule@adfc-muenchen.de",
    attachments: attachmentFiles(),
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
  // update sheet
  sheet.getRange(row, anmeldebestIndex).setValue(heuteString());
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("ADFC-RFSF")
    // .addItem("Test", "test")
    .addItem("Anmeldebestätigung senden", "anmeldebestätigung")
    .addItem("Update", "update")
    .addItem("Kursteilnehmer drucken", "printKursMembers")
    .addToUi();
}

function dispatch(e: Event) {
  let docLock = LockService.getScriptLock();
  let locked = docLock.tryLock(30000);
  if (!locked) {
    Logger.log("Could not obtain document lock");
  }
  if (!inited) init();
  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  Logger.log("dispatch sheet", sheet.getName(), range.getA1Notation());
  if (sheet.getName() == "Test") checkBuchung(e);
  if (sheet.getName() == "Buchungen") checkBuchung(e);
  if (sheet.getName() == "Email-Verifikation") verifyEmail();
  if (locked) docLock.releaseLock();
}

function verifyEmail() {
  let ssheet = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ssheet.getSheetByName("Email-Verifikation");
  if (evSheet.getLastRow() < 2) return;
  // It is a big nuisance that getSheetValues with a row count of 0 throws an error, instead of returning an empty list.
  let evalues = evSheet.getSheetValues(
    2,
    1,
    evSheet.getLastRow() - 1,
    evSheet.getLastColumn(),
  ); // Mit dieser Email-Adresse

  let numRows = buchungenSheet.getLastRow();
  if (numRows < 2) return;
  let bvalues = buchungenSheet.getSheetValues(
    2,
    1,
    numRows - 1,
    buchungenSheet.getLastColumn(),
  );
  Logger.log("bvalues %s", bvalues);

  for (let bx in bvalues) {
    let bxi = +bx; // confusingly, bx is initially a string, and is interpreted as A1Notation in sheet.getRange(bx) !
    let brow = bvalues[bxi];
    if (
      !isEmpty(brow[mailIndex - 1]) &&
      isEmpty(brow[verifikationsIndex - 1])
    ) {
      let baddr = (brow[1] as string).toLowerCase();
      for (let ex in evalues) {
        let erow = evalues[ex];
        if (erow.length < 3) continue;
        let eaddr = (erow[2] as string).toLowerCase();
        if (eaddr != baddr) continue;
        if (erow[1] != "Ja" || isEmpty(erow[2])) continue;
        // Buchungen[Verifiziert] = Email-Verif[Zeitstempel]
        buchungenSheet.getRange(bxi + 2, verifikationsIndex).setValue(erow[0]);
        brow[verifikationsIndex - 1] = erow[0];
        sendVerifEmail(brow);
        break;
      }
    }
  }
}

function sendVerifEmail(rowValues: any[]) {
  let herrFrau = rowValues[herrFrauIndex - 1];
  let name = rowValues[nameIndex - 1];
  let empfaenger = rowValues[mailIndex - 1];
  // Anrede
  let anrede: string = anredeText(herrFrau, name);
  var subject = "Emailadresse bestätigt";
  var body =
    anrede +
    ",\nvielen Dank, dass Sie Ihre E-Mail Adresse verifiziert haben.\n" +
    "In ein bis zwei Tagen bekommen Sie von uns die Bestätigung,\ndass Sie " +
    "bei dem Kurs in der Radfahrschule einen freien Platz bekommen.\n" +
    "Mit freundlichen Grüßen,\n\n" +
    "Allgemeiner Deutscher Fahrrad-Club München e.V.\n" +
    "Platenstraße 4\n" +
    "80336 München\n" +
    "Tel. 089 | 773429 Fax 089 | 778537\n" +
    "radfahrschule@adfc-muenchen.de\n" +
    "www.adfc-muenchen.de\n";
  GmailApp.sendEmail(empfaenger, subject, body);
}

function checkBuchung(e: Event) {
  let range: GoogleAppsScript.Spreadsheet.Range = e.range;
  let sheet = range.getSheet();
  let row = range.getRow();
  let cellA = range.getCell(1, 1);
  Logger.log("sheet %s row %s cellA %s", sheet, row, cellA.getA1Notation());

  if (e.namedValues["Zahlungsart"][0].startsWith("SEPA")) {
    let ibanNV = e.namedValues["Lastschrift: IBAN-Kontonummer"][0];
    let iban = ibanNV.replace(/\s/g, "").toUpperCase();
    let emailTo = e.namedValues["E-Mail-Adresse"][0].toLowerCase().trim();
    Logger.log("iban=%s emailTo=%s %s", iban, emailTo, typeof emailTo);
    if (!isValidIban(iban)) {
      sendWrongIbanEmail(anrede(e), emailTo, iban);
      cellA.setNote("Ungültige IBAN");
      return;
    }
    if (iban != ibanNV) {
      let cellIban = range.getCell(
        1,
        headers["Buchungen"]["Lastschrift: IBAN-Kontonummer"],
      );
      cellIban.setValue(iban);
    }
  }
  // Die Zellen Zustimmung und Bestätigung sind im Formular als Pflichtantwort eingetragen
  // und können garnicht anders als gesetzt sein. Sonst hier prüfen analog zu IBAN.

  let kursGebucht: string = e.namedValues[kursFrage][0];

  let msgs = [];
  let kurseS: Array<Array<string>> = kurseSheet.getSheetValues(
    2,
    1,
    kurseSheet.getLastRow(),
    kurseSheet.getLastColumn(),
  );
  let restChanged = false;
  let kursFound = false;
  for (let j = 0; j < kurseS.length; j++) {
    if (kurseS[j][0] == kursGebucht) {
      kursFound = true;
      let rest = kurseSheet.getRange(2 + j, restIndex).getValue();
      if (rest <= 0) {
        msgs.push("Der Kurs '" + kursGebucht + "' ist leider ausgebucht.");
        sheet.getRange(row, 1).setNote("Ausgebucht");
      } else {
        msgs.push("Sie sind für den Kurs '" + kursGebucht + "' vorgemerkt.");
        kurseSheet.getRange(2 + j, restIndex).setValue(rest - 1);
        restChanged = true;
      }
      break;
    }
  }
  if (!kursFound) {
    Logger.log("Kurs '" + kursGebucht + "' nicht im Kurse-Sheet!?");
  }
  if (msgs.length == 0) {
    Logger.log("keine Kurse gefunden!?");
    return;
  }
  if (restChanged) {
    updateForm();
  }
  Logger.log("msgs: ", msgs, msgs.length);
  sendeAntwort(e, msgs, sheet, row);
}

function sendeAntwort(
  e: Event,
  msgs: Array<string>,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
) {
  let emailTo = e.namedValues["E-Mail-Adresse"][0].toLowerCase().trim();
  Logger.log("emailTo=" + emailTo);

  let templateFile = "emailVerif.html";

  // do we already know this email address?
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let evSheet = ss.getSheetByName("Email-Verifikation");
  let numRows = evSheet.getLastRow();
  let evalues =
    numRows < 2
      ? []
      : evSheet.getSheetValues(2, 1, evSheet.getLastRow() - 1, 3);
  for (let i = 0; i < evalues.length; i++) {
    // Mit dieser Email-Adresse
    if (evalues[i][2].toLowerCase().trim() === emailTo) {
      templateFile = "emailReply.html"; // yes, don't ask for verification
      sheet.getRange(row, verifikationsIndex).setValue(evalues[i][0]);
    }
  }

  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    templateFile,
  );
  template.anrede = anrede(e);
  template.msgs = msgs;
  template.verifLink =
    "https://docs.google.com/forms/d/e/1FAIpQLSeQP1TTDJom91faLmhbO45z0EoDF-ZjncuUhhzQ5Pl6trnjSA/viewform?usp=pp_url&entry.1730791681=Ja&entry.1561755994=" +
    encodeURIComponent(emailTo);

  let htmlText: string = template.evaluate().getContent();
  let subject =
    templateFile === "emailVerif.html"
      ? "Bestätigung Ihrer Email-Adresse"
      : "Bestätigung Ihrer Anmeldung";
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "Radfahrschule ADFC München e.V.",
    replyTo: "radfahrschule@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
}

function anrede(e: Event) {
  // if Name is not set, nv["Name"] has value [""], i.e. not null, not [], not [null]!
  let anrede: string = e.namedValues["Anrede"][0];
  // let vorname: string = e.namedValues["Vorname"][0];
  let name: string = e.namedValues["Name"][0];

  if (anrede == "Herr") {
    anrede = "Sehr geehrter Herr ";
  } else {
    anrede = "Sehr geehrte Frau ";
  }
  Logger.log("anrede", anrede, name);
  return anrede + name;
}

function update() {
  let docLock = LockService.getScriptLock();
  let locked = docLock.tryLock(30000);
  if (!locked) {
    SpreadsheetApp.getUi().alert("Konnte Dokument nicht locken");
    return;
  }
  if (!inited) init();
  verifyEmail();
  updateReste();
  updateForm();
  docLock.releaseLock();
}

function updateReste() {
  let kurseRows = kurseSheet.getLastRow() - 1; // first row = headers
  let kurseCols = kurseSheet.getLastColumn();
  let kurseVals = kurseSheet.getRange(2, 1, kurseRows, kurseCols).getValues();
  let kurseNotes = kurseSheet.getRange(2, 1, kurseRows, kurseCols).getNotes();

  let buchungenRows = buchungenSheet.getLastRow() - 1; // first row = headers
  let buchungenCols = buchungenSheet.getLastColumn();
  let buchungenVals: any[][];
  let buchungenNotes: string[][];
  // getRange with 0 rows throws an exception instead of returning an empty array
  if (buchungenRows == 0) {
    buchungenVals = [];
    buchungenNotes = [];
  } else {
    buchungenVals = buchungenSheet
      .getRange(2, 1, buchungenRows, buchungenCols)
      .getValues();
    buchungenNotes = buchungenSheet
      .getRange(2, 1, buchungenRows, buchungenCols)
      .getNotes();
  }

  let gebuchtMap: MapS2I = {};
  for (let b = 0; b < buchungenRows; b++) {
    if (!isEmpty(buchungenNotes[b][0])) continue;
    let kurs = buchungenVals[b][kursIndexB - 1];
    let anzahl: number = gebuchtMap[kurs];
    if (anzahl == null) {
      gebuchtMap[kurs] = 1;
    } else {
      gebuchtMap[kurs] = anzahl + 1;
    }
  }

  for (let r = 0; r < kurseRows; r++) {
    if (!isEmpty(kurseNotes[r][0])) continue;
    let kurs = kurseVals[r][kursNameIndex - 1];
    let anzahl: number = kurseVals[r][anzahlIndex - 1];
    let gebucht: number = gebuchtMap[kurs];
    if (gebucht == null) gebucht = 0;
    let rest: number = anzahl - gebucht;
    if (rest < 0) {
      SpreadsheetApp.getUi().alert("Der Kurs '" + kurs + "' ist überbucht!");
      rest = 0;
    }
    let restR: number = kurseVals[r][restIndex - 1];
    if (rest !== restR) {
      kurseSheet.getRange(2 + r, restIndex).setValue(rest);
      SpreadsheetApp.getUi().alert(
        "Freie Plätze des Kurses '" +
          kurs +
          "' von " +
          restR +
          " auf " +
          rest +
          " geändert!",
      );
    }
  }
}

function updateForm() {
  let kurseHdrs = headers["Kurse"];
  let kurseRows = kurseSheet.getLastRow() - 1; // first row = headers
  let kurseCols = kurseSheet.getLastColumn();
  let kurseVals = kurseSheet.getRange(2, 1, kurseRows, kurseCols).getValues();
  let kurseNotes = kurseSheet.getRange(2, 1, kurseRows, kurseCols).getNotes();
  // Logger.log("kurse %s %s", kurseVals.length, kurseVals);
  let kurseObjs = [];
  for (let i = 0; i < kurseVals.length; i++) {
    if (!isEmpty(kurseNotes[i][0])) continue;
    let kurseObj: MapS2S = {};
    for (let hdr in kurseHdrs) {
      let idx = kurseHdrs[hdr];
      // Logger.log("hdr %s %s", hdr, idx);
      kurseObj[hdr] = kurseVals[i][idx - 1];
    }
    let ok = true;
    // check if all cells of Kurse row are nonempty
    for (let hdr in kurseHdrs) {
      if (!hdr.startsWith("Tag") && isEmpty(kurseObj[hdr])) ok = false;
    }
    // if (ok) {
    //   ok = +kurseObj["DZ-Rest"] > 0 || +kurseObj["EZ-Rest"] > 0;
    // }
    if (ok) kurseObjs.push(kurseObj);
  }
  Logger.log("kurseObjs=%s", kurseObjs);

  let formUrl = buchungenSheet.getFormUrl();
  // Logger.log("formUrl2 %s", formUrl);
  let form: GoogleAppsScript.Forms.Form = FormApp.openByUrl(formUrl);
  let items = form.getItems();
  let kurseItem: GoogleAppsScript.Forms.MultipleChoiceItem = null;
  for (let item of items) {
    //   let itemType = item.getType();
    //   Logger.log("title %s it %s %s", item.getTitle(), itemType, item.getIndex());
    if (item.getTitle() === kursFrage) {
      kurseItem = item.asMultipleChoiceItem();
      break;
    }
  }
  if (kurseItem == null) {
    SpreadsheetApp.getUi().alert("Das Formular hat keine Frage: " + kursFrage);
    return;
  }
  let choices = [];
  let descs = [];
  for (let type of [
    "Grundkurse",
    "Aufbaukurse",
    "Pedelec-Training für Senioren",
  ]) {
    descs.push(type + ":");
    for (let kursObj of kurseObjs) {
      let mr: string = kursObj["Kursname"];
      if (!mr.endsWith(type[0])) continue;

      let rest: number = +kursObj["Restplätze"];
      let freiText: string;
      if (rest <= 0) freiText = ", ausgebucht";
      else if (rest === 1) freiText = ", noch 1 Platz frei";
      else freiText = ", noch " + rest + " Plätze frei";

      let desc =
        mr +
        ", " +
        any2Str(kursObj["Datum"]) +
        ", " +
        kursObj["Uhrzeit"] +
        ", Kursort:" +
        any2Str(kursObj["Kursort"]) +
        ", Ersatztermin: " +
        any2Str(kursObj["Ersatztermin"]) +
        freiText;
      Logger.log("mr %s desc %s", mr, desc);
      descs.push(desc);
      let ok = +kursObj["Restplätze"] > 0;
      if (ok) {
        let choice = kurseItem.createChoice(mr);
        choices.push(choice);
      }
    }
  }
  let beschreibung: string;
  if (choices.length === 0) {
    beschreibung = "Leider sind alle Kurse ausgebucht!\n" + descs.join("\n");
  } else {
    beschreibung =
      "Wählen Sie einen Kurs.\nBitte beachten Sie die Anzahl noch freier Plätze!\n" +
      descs.join("\n");
  }
  kurseItem.setHelpText(beschreibung);
  kurseItem.setChoices(choices);
}

function sendWrongIbanEmail(anrede: string, empfaenger: string, iban: string) {
  var subject = "Falsche IBAN";
  var body =
    anrede +
    ",\nDie von Ihnen bei der Buchung von ADFC Mehrtageskurse übermittelte IBAN " +
    iban +
    " ist leider falsch! Bitte wiederholen Sie die Buchung mit einer korrekten IBAN.";

  body =
    body +
    "\nMit freundlichen Grüßen,\n\n" +
    "Allgemeiner Deutscher Fahrrad-Club München e.V.\n" +
    "Platenstraße 4\n" +
    "80336 München\n" +
    "Tel. 089 | 773429 Fax 089 | 778537\n" +
    "radfahrschule@adfc-muenchen.de\n" +
    "www.adfc-muenchen.de\n";

  GmailApp.sendEmail(empfaenger, subject, body);
}

let ibanLen: MapS2I = {
  NO: 15,
  BE: 16,
  DK: 18,
  FI: 18,
  FO: 18,
  GL: 18,
  NL: 18,
  MK: 19,
  SI: 19,
  AT: 20,
  BA: 20,
  EE: 20,
  KZ: 20,
  LT: 20,
  LU: 20,
  CR: 21,
  CH: 21,
  HR: 21,
  LI: 21,
  LV: 21,
  BG: 22,
  BH: 22,
  DE: 22,
  GB: 22,
  GE: 22,
  IE: 22,
  ME: 22,
  RS: 22,
  AE: 23,
  GI: 23,
  IL: 23,
  AD: 24,
  CZ: 24,
  ES: 24,
  MD: 24,
  PK: 24,
  RO: 24,
  SA: 24,
  SE: 24,
  SK: 24,
  VG: 24,
  TN: 24,
  PT: 25,
  IS: 26,
  TR: 26,
  FR: 27,
  GR: 27,
  IT: 27,
  MC: 27,
  MR: 27,
  SM: 27,
  AL: 28,
  AZ: 28,
  CY: 28,
  DO: 28,
  GT: 28,
  HU: 28,
  LB: 28,
  PL: 28,
  BR: 29,
  PS: 29,
  KW: 30,
  MU: 30,
  MT: 31,
};

function isValidIban(iban: string) {
  if (!iban.match(/^[\dA-Z]+$/)) return false;
  let len = iban.length;
  if (len != ibanLen[iban.substr(0, 2)]) return false;
  iban = iban.substr(4) + iban.substr(0, 4);
  let s = "";
  for (let i = 0; i < len; i += 1) s += parseInt(iban.charAt(i), 36);
  let m = +s.substr(0, 15) % 97;
  s = s.substr(15);
  for (; s; s = s.substr(13)) m = +("" + m + s.substr(0, 13)) % 97;
  return m == 1;
}

// I need any2str because a date copied to temp sheet showed as date.toString().
// A ' in front of the date came too late.
function any2Str(
  val: any,
  fmt: string = "E dd.MM",
  short: boolean = true,
): string {
  if (typeof val == "object" && "getUTCHours" in val) {
    let d = Utilities.formatDate(
      val,
      SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
      fmt, // "dd.MM.YYYY", "E dd.MM."
    );
    if (short) {
      d = d
        .replace("Mon", "Mo")
        .replace("Tue", "Di")
        .replace("Wed", "Mi")
        .replace("Thu", "Do")
        .replace("Fri", "Fr")
        .replace("Sat", "Sa")
        .replace("Sun", "So");
    } else {
      d = d
        .replace("Mon", "Montag")
        .replace("Tue", "Dienstag")
        .replace("Wed", "Mittwoch")
        .replace("Thu", "Donnerstag")
        .replace("Fri", "Freitag")
        .replace("Sat", "Samstag")
        .replace("Sun", "Sonntag");
    }
    return d;
  }
  return val.toString();
}

function printKursMembers() {
  Logger.log("printKursMembers");
  if (!inited) init();
  let sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Kurse") {
    SpreadsheetApp.getUi().alert(
      "Bitte eine Zeile im Sheet 'Kurse' selektieren",
    );
    return;
  }
  let curCell = sheet.getSelection().getCurrentCell();
  if (!curCell) {
    SpreadsheetApp.getUi().alert(
      "Bitte zuerst eine Zeile im Sheet 'Kurse' selektieren",
    );
    return;
  }
  let row = curCell.getRow();
  if (row < 2 || row > sheet.getLastRow()) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile ist ungültig, bitte zuerst Kurse-Zeile selektieren",
    );
    return;
  }
  let rowValues = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  let rowNote = sheet.getRange(row, 1).getNote();
  if (!isEmpty(rowNote)) {
    SpreadsheetApp.getUi().alert(
      "Die ausgewählte Zeile hat eine Notiz und ist deshalb ungültig",
    );
    return;
  }
  let kurs: string = rowValues[kursNameIndex - 1];

  let buchungenRows = buchungenSheet.getLastRow() - 1; // first row = headers
  let buchungenCols = buchungenSheet.getLastColumn();
  let buchungenVals: any[][];
  let buchungenNotes: string[][];
  // getRange with 0 rows throws an exception instead of returning an empty array
  if (buchungenRows < 1) {
    SpreadsheetApp.getUi().alert("Keine Buchungen gefunden");
    return;
  }
  buchungenVals = buchungenSheet
    .getRange(2, 1, buchungenRows, buchungenCols)
    .getValues();
  buchungenNotes = buchungenSheet.getRange(2, 1, buchungenRows, 1).getNotes();

  let ss = SpreadsheetApp.getActiveSpreadsheet();

  sheet = ss.getSheetByName("Teilnehmer-" + kurs);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Teilnehmer-" + kurs);

  let bHdrs = headers["Buchungen"];
  // first row of temp sheet: the headers
  {
    let row: string[] = [];
    for (let [_, v] of printCols) {
      row.push(v);
    }
    sheet.appendRow(row);
  }

  let rows: string[][] = [];
  for (let b = 0; b < buchungenRows; b++) {
    if (!isEmpty(buchungenNotes[b][0])) continue;
    let brow = buchungenVals[b];
    if (brow[kursIndexB - 1] !== kurs) continue;
    let row: string[] = [];
    for (let [k, _] of printCols) {
      //for the ' see https://stackoverflow.com/questions/13758913/format-a-google-sheets-cell-in-plaintext-via-apps-script
      // otherwise, telefon number 089... is printed as 89
      let val = any2Str(brow[bHdrs[k] - 1], "dd.MM");
      row.push("'" + val);
    }
    rows.push(row);
  }

  for (let row of rows) sheet.appendRow(row);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  let range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  sheet.setActiveSelection(range);
  printSelectedRange(kurs);
  //Utilities.sleep(10000);
  //ss.deleteSheet(sheet);
}

function objectToQueryString(obj: any) {
  return Object.keys(obj)
    .map(function (key) {
      return Utilities.formatString("&%s=%s", key, obj[key]);
    })
    .join("");
}

// see https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
let PRINT_OPTIONS = {
  size: 7, // paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B
  fzr: false, // repeat row headers
  portrait: true, // false=landscape
  fitw: true, // fit window or actual size
  gridlines: true, // show gridlines
  printtitle: true,
  sheetnames: true,
  pagenum: "UNDEFINED", // CENTER = show page numbers / UNDEFINED = do not show
  attachment: false,
};

let PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function printSelectedRange(kurs: string) {
  SpreadsheetApp.flush();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let range = sheet.getActiveRange();

  let gid = sheet.getSheetId();
  let printRange = objectToQueryString({
    c1: range.getColumn() - 1,
    r1: range.getRow() - 1,
    c2: range.getColumn() + range.getWidth() - 1,
    r2: range.getRow() + range.getHeight() - 1,
  });
  let url = ss.getUrl();
  let x = url.indexOf("/edit?");
  url = url.slice(0, x);
  url = url + "/export?format=pdf" + PDF_OPTS + printRange + "&gid=" + gid;

  let params: any = {
    method: "GET",
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
  };

  let resp = UrlFetchApp.fetch(url, params);
  let blob = resp.getBlob();
  blob.setName("Teilnehmer-" + kurs + ".pdf");
  let f = DriveApp.createFile(blob);
  Logger.log("file %s %s", f, f.getName());

  let htmlTemplate = HtmlService.createTemplateFromFile("print.html");
  htmlTemplate.url = url;

  let ev = htmlTemplate.evaluate();
  Logger.log("ev2" + ev.getContent());

  SpreadsheetApp.getUi().showModalDialog(
    ev.setHeight(10).setWidth(100),
    "Drucke Auswahl",
  );
}
