function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    { name: 'Crea Note', functionName: 'creaNote' },
  ];
  spreadsheet.addMenu('Crea Note azienda', menuItems);
}

function creaVariabileColonna(x) { return "B" + x; }

function creaLinkDocumento(id) {
  return 'https://docs.google.com/document/d/' + id + '/edit';
}



function creaNote() {

  // APRO LA CARTELLA IN CUI LI SALVERÒ 
  // https://drive.google.com/drive/folders/1ndeK9AHK_mCGoLUvSwrQI_p92m-FM0E_
  var folder = DriveApp.getFolderById("1ndeK9AHK_mCGoLUvSwrQI_p92m-FM0E_");

  // FOGLIO DA UTILIZZARE
  var mese = "08_22"; // SOSTITUIBILE (nome del foglio da qui prendere i dati (es.mese))

  // CONTATORI E UTILITY
  var templateID = "1nsvG-cZWoK1j2L5zJi7yZPvp6hdt8NtuBQyXkjAHMoA"; // SOSTITUIBILE (ID del foglio da utilizzare come template)
  var sheet = SpreadsheetApp.getActive().getSheetByName(mese);
  var nRighe = 2; // SOSTITUIBILE (RIGA INIZIALE)


  while (sheet.getRange(creaVariabileColonna(nRighe)).getValue() != '') {

    // RESET UTILITY
    var found = false;
    // INSERISCO QUI I FILES FRESCHI OGNI VOLTA
    var files = folder.getFiles();
    // NOME DA CERCARE
    name = sheet.getRange(creaVariabileColonna(nRighe)).getValue() + '_Note Mensili';

    // LOOP DI RICERCA DEL FILE
    while (files.hasNext() && found === false) {

      var file = files.next();
      if (file.getName() === name) {

        // AGGIORNO TROVATO
        found = true;
        Logger.log("Sono in true" + creaVariabileColonna(nRighe));

        //AGGIORNO IL LINK
        var file_trovato_ID = file.getId();

        // INSERIRE IL LINK
        var richValue = SpreadsheetApp.newRichTextValue()
          .setText(sheet.getRange(creaVariabileColonna(nRighe)).getValue())
          .setLinkUrl(creaLinkDocumento(file_trovato_ID))
          .build();
        sheet.getRange(creaVariabileColonna(nRighe)).setRichTextValue(richValue);

        //AGGIORNARE IL CONTENUTO della NOTA
        //COPIARE TUTTO IL CONTENUTO
        var doc = DocumentApp.openById(file_trovato_ID);
        var text = doc.getText();
        //CREARE NOTA
        sheet.getRange(creaVariabileColonna(nRighe)).setNote(text);
        Logger.log('added note to:');
        Logger.log(file.getName());

      }
    }

    if (found === false) {
    Logger.log("Sono in false" + creaVariabileColonna(nRighe));
    Logger.log("Creo" + name);

    // CREO
    var file_tmp = DriveApp.getFileById(templateID).makeCopy(name, folder);
    var file_tmp_id = file_tmp.getId();

    // INSERIRE IL LINK
    var richValue = SpreadsheetApp.newRichTextValue()
      .setText(sheet.getRange(creaVariabileColonna(nRighe)).getValue())
      .setLinkUrl(creaLinkDocumento(file_tmp_id))
      .build();
    sheet.getRange(creaVariabileColonna(nRighe)).setRichTextValue(richValue);
    }

    nRighe += 1;
  }



}