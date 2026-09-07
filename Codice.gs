function doGet() {
  const faviconUrl = 'https://i.postimg.cc/fT1WP57j/FITNESS-CON-VALENTINA-icon.png';
  
  // Recupera l'email dell'utente attualmente loggato su Google
  const emailUtente = Session.getActiveUser().getEmail().toLowerCase();
  
  // Imposta le email autorizzate separate da ||
  const isIstruttore = (emailUtente === "salvatore.vatrella@gmail.com" || emailUtente === "valentina.varricchio@gmail.com");

  // Crea il template e passa la variabile booleana alla pagina HTML
  const template = HtmlService.createTemplateFromFile('Index');
  template.isIstruttore = isIstruttore;

  return template
    .evaluate()
    .setTitle('Fitness con Valentina')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setFaviconUrl(faviconUrl)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// SEZIONE: INIZIALIZZAZIONE FOGLI E STRUTTURE
// ==========================================
function verificaInizializzazioneFogli() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Utenti: Nome, Cognome, Numero telefono, E-mail, Acquistate, Utilizzate, Da Saldare
  let sUtenti = ss.getSheetByName('Utenti');
  if (!sUtenti) {
    sUtenti = ss.insertSheet('Utenti');
    sUtenti.appendRow(['Nome', 'Cognome', 'Numero telefono', 'E-mail', 'Acquistate', 'Utilizzate', 'Da Saldare']);
  }

  // 2. Acquisti: Nome, Cognome, Categoria, Lezioni Acquistate, Lezioni Utilizzate, Da Saldare
  let sAcquisti = ss.getSheetByName('Acquisti');
  if (!sAcquisti) {
    sAcquisti = ss.insertSheet('Acquisti');
    sAcquisti.appendRow(['Nome', 'Cognome', 'Categoria', 'Acquistate', 'Utilizzate', 'Da Saldare']);
  }

  // 3. Listini: Categoria, Num Pacchetto 1, Prezzo 1, Num Pacchetto 2, Prezzo 2, Num Pacchetto 3, Prezzo 3
  let sListini = ss.getSheetByName('Listini');
  if (!sListini) {
    sListini = ss.insertSheet('Listini');
    sListini.appendRow(['Categoria', 'Num Pacchetto 1', 'Prezzo 1', 'Num Pacchetto 2', 'Prezzo 2', 'Num Pacchetto 3', 'Prezzo 3']);
    sListini.appendRow(['Pilates', 1, 15, 5, 65, 10, 120]);
    sListini.appendRow(['Yoga', 1, 15, 5, 65, 10, 120]);
  }

// 4. Calendario: Nome, Data, Ora Inizio, Ora Fine, Luogo, Posti Massimi, Posti Prenotati, Categoria, ID Lezione (Colonna I)
 let sCal = ss.getSheetByName('Calendario');
  if (!sCal) {
    sCal = ss.insertSheet('Calendario');
    sCal.appendRow(['Nome', 'Data', 'Ora Inizio', 'Ora Fine', 'Luogo', 'Posti Massimi', 'Posti Prenotati', 'Categoria', 'ID Lezione']);
  } else {
    // Se il foglio esiste già ma non ha l'intestazione in colonna I, la impostiamo
    let headerColI = sCal.getRange(1, 9).getValue();
    if (!headerColI) {
      sCal.getRange(1, 9).setValue('ID Lezione');
    }
  }

  // 5. Prenotazioni: Nome, Cognome, Data, Ora Inizio, Ora Fine, Nome Lezione, Luogo, Categoria (Colonna H)
  let sPrenotazioni = ss.getSheetByName('Prenotazioni');
  if (!sPrenotazioni) {
    sPrenotazioni = ss.insertSheet('Prenotazioni');
    sPrenotazioni.appendRow(['Email Utente', 'Nome Utente', 'ID Lezione', 'Data Prenotazione', 'ID Prenotazione']); 
    // Dove ID Lezione va in Colonna I e ID Prenotazione in Colonna J
  }
}

// ==========================================
// SEZIONE: CALENDARIO & LEZIONI
// ==========================================

function getLezioni() {
  verificaInizializzazioneFogli();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Calendario');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const tz = ss.getSpreadsheetTimeZone();
  const lezioni = [];
  
  // Otteniamo il timestamp corrente per il confronto temporale
  const adesso = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    
    let rawDate = row[1];
    let dataIso = "";
    let dataFormatted = "";
    
    if (rawDate instanceof Date) {
      dataIso = Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd');
      dataFormatted = Utilities.formatDate(rawDate, tz, 'dd/MM/yyyy');
    } else if (rawDate) {
      dataFormatted = String(rawDate);
      const parts = dataFormatted.split('/');
      if (parts.length === 3) {
        dataIso = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
      } else {
        dataIso = dataFormatted;
      }
    }
    
    const oraInizio = row[2] instanceof Date ? Utilities.formatDate(row[2], tz, 'HH:mm') : String(row[2] || '');
    const oraFine = row[3] instanceof Date ? Utilities.formatDate(row[3], tz, 'HH:mm') : String(row[3] || '');
    const luogo = String(row[4] || '');
    const postiMassimi = Number(row[5]) || 0;
    const postiPrenotati = Number(row[6]) || 0;
    const categoria = String(row[7] || 'Generale');
    
    // Leggiamo l'ID univoco dalla Colonna I (indice 8), se non esiste lo generiamo al volo per retrocompatibilità
    let idLezione = row[8] ? String(row[8]).trim() : Utilities.getUuid();
    
    // --- CALCOLO SE LA LEZIONE È PASSATA ---
    let isPassata = false;
    if (dataIso) {
      // Costruiamo la data/ora di inizio della lezione
      const dataOraLezioneStr = oraInizio ? `${dataIso}T${oraInizio}:00` : `${dataIso}T23:59:59`;
      const dataLezioneObj = new Date(dataOraLezioneStr);
      
      // Se la data è valida, confrontiamo con il momento attuale
      if (!isNaN(dataLezioneObj.getTime())) {
        isPassata = adesso > dataLezioneObj;
      }
    }
    // --------------------------------------
    
    const mapsUrl = luogo ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(luogo)}` : '';
    
    lezioni.push({
      rowIndex: i + 1, // Manteniamo traccia della riga Excel se serve
      id: idLezione,   // ID Univoco Colonna I
      nome: String(row[0]),
      dataIso: dataIso,
      dataFormatted: dataFormatted,
      oraInizio: oraInizio,
      oraFine: oraFine,
      luogo: luogo,
      mapsUrl: mapsUrl,
      postiMassimi: postiMassimi,
      postiPrenotati: postiPrenotati,
      categoria: categoria,
      esaurito: postiPrenotati >= postiMassimi,
      passata: isPassata // <-- Nuovo flag utile per il frontend
    });
  }
  
  return lezioni;
}

function aggiungiNuovaLezione(nome, data, oraInizio, oraFine, luogo, postiMassimi, categoria) {
  verificaInizializzazioneFogli();
  nome = String(nome).trim();
  data = String(data).trim();
  oraInizio = String(oraInizio).trim();
  oraFine = String(oraFine).trim();
  luogo = String(luogo).trim();
  postiMassimi = Number(postiMassimi) || 0;
  categoria = String(categoria || 'Generale').trim();

  if (!nome || !data || !oraInizio || !oraFine || !postiMassimi) {
    return { success: false, message: "Compilare tutti i campi obbligatori." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Calendario');
  
  const parts = data.split('-');
  const dataFormatted = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : data;

  // Generiamo un ID univoco sicuro (es. UUID nativo di Google Apps Script)
  const idUnivoco = Utilities.getUuid();

  // Inseriamo i dati fino alla colonna H e l'ID univoco nella colonna I (9° elemento)
  sheet.appendRow([nome, dataFormatted, oraInizio, oraFine, luogo, postiMassimi, 0, categoria, idUnivoco]);
  
  return { success: true, message: "Lezione aggiunta con successo!", id: idUnivoco };
}

function aggiornaLezione(idLezione, nome, data, oraInizio, oraFine, luogo, postiMassimi, categoria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCalendario = ss.getSheetByName('Calendario');
  if (!sheetCalendario) return { success: false, message: "Tab 'Calendario' non trovato." };

  // 1. Cerchiamo la riga nel Calendario confrontando l'ID Univoco nella Colonna I (indice 8)
  const dataCalendario = sheetCalendario.getDataRange().getValues();
  let rigaTrovata = -1;
  
  const idCercato = String(idLezione || "").trim();

  for (let i = 1; i < dataCalendario.length; i++) {
    const currentId = String(dataCalendario[i][8] || "").trim(); // Colonna I
    if (currentId === idCercato) {
      rigaTrovata = i + 1; // Le righe di Sheets partono da 1
      break;
    }
  }

  if (rigaTrovata === -1) {
    return { success: false, message: "Lezione non trovata tramite ID univoco." };
  }

  // 2. Formatta la nuova data nel formato dd/MM/yyyy
  const parts = data.split('-');
  const dataFormatted = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : data;

  // 3. Aggiorna i dati nel foglio 'Calendario'
  sheetCalendario.getRange(rigaTrovata, 1).setValue(nome);          // A: Nome lezione
  sheetCalendario.getRange(rigaTrovata, 2).setValue(dataFormatted);   // B: Data Lezione
  sheetCalendario.getRange(rigaTrovata, 3).setValue(oraInizio);       // C: Ora inizio
  sheetCalendario.getRange(rigaTrovata, 4).setValue(oraFine);         // D: ora fine
  sheetCalendario.getRange(rigaTrovata, 5).setValue(luogo);           // E: Luogo
  sheetCalendario.getRange(rigaTrovata, 6).setValue(postiMassimi);    // F: Posti max
  // La colonna G (7) è "Posti Prenotati" e non viene toccata qui
  sheetCalendario.getRange(rigaTrovata, 8).setValue(categoria);       // H: Categoria
  // La colonna I (9) mantiene l'idLezione immutato

  // 4. Aggiorna le prenotazioni nel foglio 'Prenotazioni' basandosi sull'ID Lezione (Colonna I / indice 8 nel foglio Prenotazioni)
  const sheetPrenotazioni = ss.getSheetByName('Prenotazioni');
  if (sheetPrenotazioni) {
    const dataPrenotazioni = sheetPrenotazioni.getDataRange().getValues();
    
    for (let j = 1; j < dataPrenotazioni.length; j++) {
      let idLezionePrenotata = String(dataPrenotazioni[j][8] || "").trim(); // Supponendo che l'ID Lezione sia in Colonna I (indice 8) del foglio Prenotazioni

      if (idLezionePrenotata === idCercato) {
        let rigaPrenotazione = j + 1;
        
        // Aggiorna i campi descrittivi della lezione duplicati nel foglio prenotazioni (se presenti)
        sheetPrenotazioni.getRange(rigaPrenotazione, 3).setValue(dataFormatted); // Data lezione (se in colonna C)
        sheetPrenotazioni.getRange(rigaPrenotazione, 4).setValue(oraInizio);    // Ora inizio
        sheetPrenotazioni.getRange(rigaPrenotazione, 5).setValue(oraFine);      // Ora fine
        sheetPrenotazioni.getRange(rigaPrenotazione, 6).setValue(nome);         // Nome Lezione
        sheetPrenotazioni.getRange(rigaPrenotazione, 7).setValue(luogo);         // Luogo
        sheetPrenotazioni.getRange(rigaPrenotazione, 8).setValue(categoria);    // Categoria
      }
    }
  }

  return { success: true, message: "Lezione e relative prenotazioni aggiornate con successo!" };
}

function eliminaLezione(idLezione) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCalendario = ss.getSheetByName('Calendario');
  if (!sheetCalendario) return { success: false, message: "Tab 'Calendario' non trovato." };

  const data = sheetCalendario.getDataRange().getValues();
  let rigaTrovata = -1;
  const idCercato = String(idLezione || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][8] || "").trim() === idCercato) { // Colonna I (indice 8)
      rigaTrovata = i + 1;
      break;
    }
  }

  if (rigaTrovata === -1) {
    return { success: false, message: "Lezione non trovata." };
  }

  // Elimina la riga dal foglio Calendario
  sheetCalendario.deleteRow(rigaTrovata);

  // (Facoltativo ma consigliato) Elimina o pulisci anche le prenotazioni collegate nel foglio 'Prenotazioni'
  const sheetPrenotazioni = ss.getSheetByName('Prenotazioni');
  if (sheetPrenotazioni) {
    const dataPrenotazioni = sheetPrenotazioni.getDataRange().getValues();
    // Scorri al contrario per eliminare le righe senza alterare gli indici
    for (let j = dataPrenotazioni.length - 1; j >= 1; j--) {
      if (String(dataPrenotazioni[j][8] || "").trim() === idCercato) { // Supponendo ID Lezione in Colonna I del foglio Prenotazioni
        sheetPrenotazioni.deleteRow(j + 1);
      }
    }
  }

  return { success: true, message: "Lezione eliminata con successo." };
}

// ==========================================
// SEZIONE: PRENOTAZIONI & CREDITI PER CATEGORIA
// ==========================================

function prenotaLezione(idLezione, telefono) {
  verificaInizializzazioneFogli();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const telCercato = String(telefono || "").trim();
  if (!telCercato) {
    return { success: false, message: "Inserisci un numero di telefono valido." };
  }
  
  // 1. Verifichiamo la lezione nel foglio Calendario tramite l'ID Lezione
  const sheetCal = ss.getSheetByName('Calendario');
  if (!sheetCal) return { success: false, message: "Foglio Calendario non trovato." };
  
  const dataCal = sheetCal.getDataRange().getValues();
  let rigaLezione = -1;
  let datiLezione = {};
  
  for (let i = 1; i < dataCal.length; i++) {
    const currentIdLezione = String(dataCal[i][8] || "").trim(); // ID Lezione in Colonna I (indice 8)
    if (currentIdLezione === String(idLezione).trim()) {
      rigaLezione = i + 1;
      
      datiLezione = {
        nomeLezione: sheetCal.getRange(rigaLezione, 1).getValue(),   // Colonna A: Nome Lezione
        dataFormatted: sheetCal.getRange(rigaLezione, 2).getValue(), // Colonna C: Data lezione
        oraInizio: sheetCal.getRange(rigaLezione, 3).getValue(),     // Colonna D: Ora inizio
        oraFine: sheetCal.getRange(rigaLezione, 4).getValue(),       // Colonna E: Ora fine
        luogo: sheetCal.getRange(rigaLezione, 5).getValue(),         // Colonna G: Luogo
        categoria: String(sheetCal.getRange(rigaLezione, 8).getValue()).trim(), // Colonna H: Categoria
        idLezione: currentIdLezione
      };
      break;
    }
  }
  
  if (rigaLezione === -1) {
    return { success: false, message: "Lezione non trovata o ID non valido." };
  }
  
  // 2. Verifichiamo l'utente nel foglio "Utenti" tramite il numero di telefono
  const sheetUtenti = ss.getSheetByName('Utenti');
  if (!sheetUtenti) {
    return { success: false, message: "Foglio Utenti non trovato." };
  }
  
  const dataUtenti = sheetUtenti.getDataRange().getValues();
  let utenteTrovato = null;
  
  for (let i = 1; i < dataUtenti.length; i++) {
    const rigaStr = dataUtenti[i].map(cell => String(cell).trim());
    if (rigaStr.includes(telCercato)) {
      utenteTrovato = {
        nome: dataUtenti[i][0],    // Nome
        cognome: dataUtenti[i][1], // Cognome
        telefono: telCercato
      };
      break;
    }
  }
  
  if (!utenteTrovato) {
    return { 
      success: false, 
      code: "UTENTE_NON_REGISTRATO", 
      message: "Numero di telefono non registrato nel sistema." 
    };
  }
  
  // 3. CONTROLLO ANTI-DOPPIA PRENOTAZIONE
  const sheetPrenotazioni = ss.getSheetByName('Prenotazioni') || ss.insertSheet('Prenotazioni');
  const dataPrenotazioni = sheetPrenotazioni.getDataRange().getValues();
  
  for (let i = 1; i < dataPrenotazioni.length; i++) {
    const pIdLezione = String(dataPrenotazioni[i][8] || "").trim(); // Colonna I: ID Lezione
    const pTelefono = String(dataPrenotazioni[i][10] || "").trim();  // Colonna K: Telefono
    
    if (pTelefono === telCercato && pIdLezione === String(idLezione).trim()) {
      return { success: false, message: "Hai già prenotato questa lezione!" };
    }
  }
  
  // 4. CONTROLLO E AGGIORNAMENTO CREDITI NEL FOGLIO "Acquisti"
  const sheetAcquisti = ss.getSheetByName('Acquisti');
  let creditoTrovato = false;
  let rigaAcquistoTrovata = -1;
  
  if (sheetAcquisti) {
    const dataAcquisti = sheetAcquisti.getDataRange().getValues();
    for (let i = 1; i < dataAcquisti.length; i++) {
      const acqTel = String(dataAcquisti[i][7] || "").trim();      // Colonna H: Telefono (indice 7)
      const acqCategoria = String(dataAcquisti[i][3] || "").trim();// Colonna D: Tipocategoria (indice 3)
      const lezioniAcquistate = Number(dataAcquisti[i][4]) || 0;   // Colonna E: Lezioni Acquistate (indice 4)
      let lezioniUtilizzate = Number(dataAcquisti[i][5]) || 0;      // Colonna F: Lezioni Utilizzate (indice 5)
      
      let lezioniResidue = lezioniAcquistate - lezioniUtilizzate;
      
      // Se troviamo il pacchetto dell'utente per questa categoria con lezioni residue disponibili > 0
      if (acqTel === telCercato && acqCategoria.toLowerCase() === datiLezione.categoria.toLowerCase() && lezioniResidue > 0) {
        rigaAcquistoTrovata = i + 1;
        creditoTrovato = true;
        
        // Incrementiamo di 1 le lezioni utilizzate nel foglio Acquisti (Colonna F)
        sheetAcquisti.getRange(rigaAcquistoTrovata, 6).setValue(lezioniUtilizzate + 1);
        break;
      }
    }
  }
  
  if (sheetAcquisti && !creditoTrovato) {
    return { 
      success: false, 
      code: "CREDITI_ESAURITI", 
      categoria: datiLezione.categoria,
      telefono: telCercato,
      message: "Non hai crediti disponibili per questa categoria." 
    };
  }
  
  // 5. Controlliamo i posti nel Calendario
  let postiMassimi = Number(sheetCal.getRange(rigaLezione, 6).getValue()) || 0;
  let postiPrenotati = Number(sheetCal.getRange(rigaLezione, 7).getValue()) || 0;
  
  if (postiPrenotati >= postiMassimi) {
    // Se la lezione è piena, ripristiniamo le lezioni utilizzate scalate in precedenza
    if (rigaAcquistoTrovata !== -1) {
      let currentUtilizzate = Number(sheetAcquisti.getRange(rigaAcquistoTrovata, 6).getValue()) || 0;
      if (currentUtilizzate > 0) {
        sheetAcquisti.getRange(rigaAcquistoTrovata, 6).setValue(currentUtilizzate - 1);
      }
    }
    return { success: false, message: "Spiacenti, la lezione è al completo." };
  }
  
  // 6. Incrementiamo i posti prenotati nel Calendario
  sheetCal.getRange(rigaLezione, 7).setValue(postiPrenotati + 1);
  
  // 7. Registriamo la prenotazione nel foglio "Prenotazioni" (da A a K)
  if (sheetPrenotazioni.getLastRow() === 0) {
    sheetPrenotazioni.getRange(1, 1, 1, 11).setValues([[
      'Nome', 'Cognome', 'Data lezione', 'ora inizio', 'ora fine', 
      'Nome Lezione', 'Luogo', 'Categoria', 'ID Lezione', 'ID Prenotazione', 'Telefono'
    ]]);
  }
  
  const idPrenotazione = Utilities.getUuid();
  const ultimaRigaPrenotazioni = Math.max(2, sheetPrenotazioni.getLastRow() + 1);
  
  const rigaDaInserire = [
    utenteTrovato.nome,       // A: Nome
    utenteTrovato.cognome,    // B: Cognome
    datiLezione.dataFormatted,// C: Data lezione
    datiLezione.oraInizio,    // D: ora inizio
    datiLezione.oraFine,      // E: ora fine
    datiLezione.nomeLezione,  // F: Nome Lezione
    datiLezione.luogo,        // G: Luogo
    datiLezione.categoria,    // H: Categoria
    idLezione,                // I: ID Lezione
    idPrenotazione,           // J: ID Prenotazione
    telCercato                // K: Telefono
  ];
  
  sheetPrenotazioni.getRange(ultimaRigaPrenotazioni, 1, 1, rigaDaInserire.length).setValues([rigaDaInserire]);
  
  return { 
    success: true, 
    message: "Prenotazione effettuata con successo!", 
    idPrenotazione: idPrenotazione 
  };
}

function aggiornaTotaliUtente(nome, cognome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUtenti = ss.getSheetByName('Utenti');
  const sheetAcquisti = ss.getSheetByName('Acquisti');
  
  const acqData = sheetAcquisti.getDataRange().getValues();
  let totAcquistate = 0;
  let totUtilizzate = 0;

  for (let i = 1; i < acqData.length; i++) {
    if (String(acqData[i][0]).trim().toLowerCase() === nome.toLowerCase() && String(acqData[i][1]).trim().toLowerCase() === cognome.toLowerCase()) {
      totAcquistate += Number(acqData[i][3]) || 0;
      totUtilizzate += Number(acqData[i][4]) || 0;
    }
  }

  const utentiData = sheetUtenti.getDataRange().getValues();
  for (let i = 1; i < utentiData.length; i++) {
    if (String(utentiData[i][0]).trim().toLowerCase() === nome.toLowerCase() && String(utentiData[i][1]).trim().toLowerCase() === cognome.toLowerCase()) {
      sheetUtenti.getRange(i + 1, 5).setValue(totAcquistate);
      sheetUtenti.getRange(i + 1, 6).setValue(totUtilizzate);
      break;
    }
  }
}

function aggiornaUtenteCompleto(datiModifica) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUtenti = ss.getSheetByName('Utenti');
    const sheetAcquisti = ss.getSheetByName('Acquisti'); 
    
    if (!sheetUtenti) return { success: false, message: "Foglio Utenti non trovato." };
    
    const rowIndex = Number(datiModifica.rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: "Riga utente non valida." };
    
    const telefonoUtente = String(datiModifica.telefono || "").trim();

    // 1. Aggiorna i dati anagrafici nel foglio Utenti
    sheetUtenti.getRange(rowIndex, 1).setValue(datiModifica.nome);
    sheetUtenti.getRange(rowIndex, 2).setValue(datiModifica.cognome);
    sheetUtenti.getRange(rowIndex, 3).setValue(telefonoUtente);
    sheetUtenti.getRange(rowIndex, 4).setValue(datiModifica.email || "");
    
    // 2. Gestione dei pacchetti/categorie
    if (datiModifica.pacchetti && Array.isArray(datiModifica.pacchetti) && sheetAcquisti) {
      const acqData = sheetAcquisti.getDataRange().getValues();
      
      datiModifica.pacchetti.forEach(p => {
        const catCorrente = String(p.categoria || "").trim();
        if (!catCorrente) return;
        
        let rigaTrovata = -1;
        // Cerca se esiste già una riga per questo utente e questa categoria nel foglio Acquisti
        for (let j = 1; j < acqData.length; j++) {
          const aNome = String(acqData[j][1] || "").trim().toLowerCase();
          const aCognome = String(acqData[j][2] || "").trim().toLowerCase();
          const aCat = String(acqData[j][3] || "").trim().toLowerCase();
          
          if (aNome === datiModifica.nome.toLowerCase() && 
              aCognome === datiModifica.cognome.toLowerCase() && 
              aCat === catCorrente.toLowerCase()) {
            rigaTrovata = j + 1;
            break;
          }
        }
        
        if (rigaTrovata !== -1) {
          // Aggiorna la riga esistente (inclusa la colonna H / 8 per il telefono)
          sheetAcquisti.getRange(rigaTrovata, 5).setValue(Number(p.acquistate) || 0); // Acquistate
          sheetAcquisti.getRange(rigaTrovata, 6).setValue(Number(p.utilizzate) || 0); // Utilizzate
          sheetAcquisti.getRange(rigaTrovata, 7).setValue(Number(p.daSaldare) || 0);  // Da saldare
          sheetAcquisti.getRange(rigaTrovata, 8).setValue(telefonoUtente);           // Telefono (Colonna H)
        } else {
          // Se è una categoria nuova, crea una nuova riga inserendo anche il telefono in fondo
          sheetAcquisti.appendRow([
            new Date(), 
            datiModifica.nome, 
            datiModifica.cognome, 
            catCorrente, 
            Number(p.acquistate) || 0, 
            Number(p.utilizzate) || 0, 
            Number(p.daSaldare) || 0,
            telefonoUtente // Colonna H: Telefono
          ]);
        }
      });
    }
    
    // Aggiorna i totali generali di riepilogo dell'utente
    if (typeof aggiornaTotaliUtente === 'function') {
      aggiornaTotaliUtente(datiModifica.nome, datiModifica.cognome);
    }
    
    return { success: true, message: "Utente e pacchetti aggiornati con successo!" };
  } catch (err) {
    return { success: false, message: "Errore durante l'aggiornamento: " + err.message };
  }
}

function getPrenotazioniUtente(telefono) {
  verificaInizializzazioneFogli();
  const telCercato = String(telefono || "").trim();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPrenotazioni = ss.getSheetByName('Prenotazioni');
  const sheetUtenti = ss.getSheetByName('Utenti');
  const sheetAcquisti = ss.getSheetByName('Acquisti');
  const tz = ss.getSpreadsheetTimeZone();
  
  let utenteTrovato = false;
  let riepilogoAcquisti = [];
  
  // 1. Controllo nel foglio Utenti (Telefono in Colonna C -> indice 2)
  if (sheetUtenti) {
    const utentiData = sheetUtenti.getDataRange().getValues();
    for (let i = 1; i < utentiData.length; i++) {
      const uTel = String(utentiData[i][2] || "").trim();
      if (uTel === telCercato) {
        utenteTrovato = true;
        break;
      }
    }
  }

  // 2. Controllo nel foglio Acquisti (Telefono in Colonna H -> indice 7)
  if (sheetAcquisti) {
    const acqData = sheetAcquisti.getDataRange().getValues();
    for (let i = 1; i < acqData.length; i++) {
      const aTel = String(acqData[i][7] || "").trim(); // Colonna H = indice 7
      if (aTel === telCercato) {
        riepilogoAcquisti.push({
          categoria: acqData[i][3],                       // Colonna D
          acquistate: Number(acqData[i][4]) || 0,     // Colonna E
          utilizzate: Number(acqData[i][5]) || 0,     // Colonna F
          rimanenti: Math.max(0, (Number(acqData[i][4]) || 0) - (Number(acqData[i][5]) || 0)),
          daSaldare: Number(acqData[i][6]) || 0       // Colonna G
        });
      }
    }
  }

  // 3. Recupero Prenotazioni (Telefono in Colonna K -> indice 10 e ID Prenotazione in Colonna J -> indice 9)
  const miePrenotazioni = [];
  if (sheetPrenotazioni) {
    const data = sheetPrenotazioni.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const pTel = String(data[i][10] || "").trim(); // Colonna K = indice 10
      
      if (pTel === telCercato) {
        const dataFormatted = data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], tz, 'dd/MM/yyyy') : String(data[i][2]);
        const oraInizio = data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], tz, 'HH:mm') : String(data[i][3]);
        const oraFine = data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], tz, 'HH:mm') : String(data[i][4]);
        const luogo = String(data[i][6] || '');
        const categoria = String(data[i][7] || ''); 
        const idPrenotazione = String(data[i][9] || '').trim(); // Colonna J (indice 9) -> ID univoco prenotazione
        const mapsUrl = luogo ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(luogo)}` : '';

        miePrenotazioni.push({
          rowIndex: i + 1,
          idPrenotazione: idPrenotazione,
          lezione: data[i][5],
          data: dataFormatted,
          oraInizio: oraInizio,
          oraFine: oraFine,
          luogo: luogo,
          categoria: categoria,
          mapsUrl: mapsUrl
        });
      }
    }
  }

  return { utenteTrovato: utenteTrovato, riepilogoAcquisti: riepilogoAcquisti, prenotazioni: miePrenotazioni };
}

function cancellaPrenotazione(idPrenotazione) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPrenotazioni = ss.getSheetByName('Prenotazioni');
  const sheetAcquisti = ss.getSheetByName('Acquisti');
  const sheetCalendario = ss.getSheetByName('Calendario');
  
  if (!sheetPrenotazioni) return { success: false, message: "Foglio Prenotazioni non trovato." };

  const idCercato = String(idPrenotazione || "").trim();
  if (!idCercato) {
    return { success: false, message: "Errore: ID Prenotazione non valido." };
  }

  const dataPren = sheetPrenotazioni.getDataRange().getValues();
  let rigaPrenotazioneTrovata = -1;
  let datiPrenotazione = null;

  // Cerchiamo la prenotazione basandoci sull'ID Prenotazione (Colonna J / indice 9)
  for (let i = 1; i < dataPren.length; i++) {
    const currentId = String(dataPren[i][9] || "").trim();
    if (currentId === idCercato) {
      rigaPrenotazioneTrovata = i + 1;
      datiPrenotazione = {
        nome: String(dataPren[i][0] || "").trim(),        
        cognome: String(dataPren[i][1] || "").trim(),      
        idLezione: String(dataPren[i][8] || "").trim(),    
        categoria: String(dataPren[i][7] || "").trim(),    
        telefono: String(dataPren[i][10] || "").trim()     
      };
      break;
    }
  }

  if (rigaPrenotazioneTrovata === -1 || !datiPrenotazione) {
    return { success: false, message: "Prenotazione non trovata nel sistema." };
  }

  // --- CONTROLLO DI SICUREZZA BLINDATO: LA LEZIONE È GIÀ PASSATA? ---
  if (sheetCalendario && datiPrenotazione.idLezione) {
    const calData = sheetCalendario.getDataRange().getValues();
    let dataLezioneTrovata = null;
    let oraLezioneTrovata = null;

    // L'ID della lezione nel foglio Calendario si trova nella Colonna I (Indice 8)
    for (let k = 1; k < calData.length; k++) {
      if (String(calData[k][8] || "").trim() === datiPrenotazione.idLezione) {
        dataLezioneTrovata = calData[k][1]; // Colonna B = Data
        oraLezioneTrovata = calData[k][2];  // Colonna C = Orario
        break;
      }
    }

    if (dataLezioneTrovata) {
      let dLezione;
      
      if (Object.prototype.toString.call(dataLezioneTrovata) === '[object Date]') {
        dLezione = new Date(dataLezioneTrovata.getTime());
      } else {
        dLezione = new Date(dataLezioneTrovata);
      }
      
      if (oraLezioneTrovata) {
        const partiOra = String(oraLezioneTrovata).trim().split(':');
        if (partiOra.length >= 2) {
          dLezione.setHours(parseInt(partiOra[0], 10), parseInt(partiOra[1], 10), 0, 0);
        }
      } else {
        dLezione.setHours(23, 59, 59, 999);
      }

      if (new Date() > dLezione) {
        return { success: false, message: "Impossibile cancellare: la lezione è già iniziata o passata." };
      }
    }
  }
  // -----------------------------------------------------------------

  // 1. Cancella la riga della prenotazione
  sheetPrenotazioni.deleteRow(rigaPrenotazioneTrovata);

  // 2. Rimborsa la lezione nel foglio Acquisti
  if (sheetAcquisti) {
    const acqData = sheetAcquisti.getDataRange().getValues();
    let rigaAcquistoTrovata = -1;
    let utilizzateAttuali = 0;

    const telCercato = datiPrenotazione.telefono.toLowerCase();
    const nomeCercato = datiPrenotazione.nome.toLowerCase();
    const cognomeCercato = datiPrenotazione.cognome.toLowerCase();
    const catCercata = datiPrenotazione.categoria.toLowerCase();

    if (telCercato) {
      for (let i = 1; i < acqData.length; i++) {
        const aTel = String(acqData[i][7] || "").trim().toLowerCase();      
        const aCat = String(acqData[i][3] || "").trim().toLowerCase();      
        const aUtilizzate = Number(acqData[i][5]) || 0;      

        if (aTel === telCercato && aCat === catCercata && aUtilizzate > 0) {
          rigaAcquistoTrovata = i + 1;
          utilizzateAttuali = aUtilizzate;
          break; 
        }
      }
    }

    if (rigaAcquistoTrovata === -1) {
      for (let i = 1; i < acqData.length; i++) {
        const aNome = String(acqData[i][1] || "").trim().toLowerCase();    
        const aCognome = String(acqData[i][2] || "").trim().toLowerCase(); 
        const aCat = String(acqData[i][3] || "").trim().toLowerCase();    
        const aUtilizzate = Number(acqData[i][5]) || 0;                                 

        if (aNome === nomeCercato && aCognome === cognomeCercato && aCat === catCercata && aUtilizzate > 0) {
          rigaAcquistoTrovata = i + 1;
          utilizzateAttuali = aUtilizzate;
          break;
        }
      }
    }

    if (rigaAcquistoTrovata !== -1) {
      const nuovoValoreUtilizzate = Math.max(0, utilizzateAttuali - 1);
      sheetAcquisti.getRange(rigaAcquistoTrovata, 6).setValue(nuovoValoreUtilizzate);
    }
  }

  if (typeof aggiornaTotaliUtente === 'function') {
    aggiornaTotaliUtente(datiPrenotazione.nome, datiPrenotazione.cognome);
  }

  // 3. Decrementa il contatore dei posti nel Calendario (cercando l'ID nella Colonna I / indice 8)
  if (sheetCalendario && datiPrenotazione.idLezione) {
    const calData = sheetCalendario.getDataRange().getValues();
    for (let k = 1; k < calData.length; k++) {
      if (String(calData[k][8] || "").trim() === datiPrenotazione.idLezione) {
        const pPrenotati = Number(calData[k][6]) || 0; // Colonna G (indice 6)
        if (pPrenotati > 0) {
          sheetCalendario.getRange(k + 1, 7).setValue(pPrenotati - 1);
        } else {
          sheetCalendario.getRange(k + 1, 7).setValue(0);
        }
        break;
      }
    }
  }

  return { success: true, message: "Prenotazione cancellata e credito rimborsato con successo!" };
}

// ==========================================
// SEZIONE: GESTIONE UTENTI & REGISTRAZIONE
// ==========================================

function registraNuovoUtenteInTab(nome, cognome, telefono, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Utenti") || ss.getSheetByName("Anagrafica");

  if (!sheet) {
    return { success: false, message: "⚠️ Foglio anagrafica utenti non trovato." };
  }

  // Controlliamo se l'utente esiste già
  var data = sheet.getDataRange().getValues();
  var emailClean = String(email || "").trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][3] || "").trim().toLowerCase();
    if (rowEmail === emailClean) {
      return { success: false, message: "⚠️ Attenzione: un utente con questa email è già registrato." };
    }
  }

  var dataRegistrazione = new Date();

  // Aggiungiamo i dati dell'utente nel foglio
  sheet.appendRow([nome, cognome, telefono, email, dataRegistrazione]);

  // --- INVIO EMAIL DI NOTIFICA ALL'ISTRUTTORE ---
  try {
    const tuaEmail = "salvatore.vatrella@gmail.com"; 
    if (tuaEmail) {
      const oggetto = `Nuova Registrazione Utente: ${nome} ${cognome}`;
      
      const corpoHtml = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
        </head>
        <body style="margin: 0; padding: 0; background-color: #f4f6f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" style="table-layout: fixed; background-color: #f4f6f9; padding: 30px 0;">
            <tr>
              <td align="center">
                <table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05); overflow: hidden;">
                  
                  <!-- Header con il colore viola -->
                  <tr>
                    <td style="background: linear-gradient(135deg, #FF64AE 0%, #FF58A1 100%); padding: 35px 40px; text-align: center;">
                      <h1 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: bold; letter-spacing: 0.5px;">
                        Nuovo Utente Registrato
                      </h1>
                      <p style="color: #f3e8ff; margin: 6px 0 0 0; font-size: 14px; font-weight: 500;">
                        Fitness con Valentina
                      </p>
                    </td>
                  </tr>

                  <!-- Corpo del messaggio -->
                  <tr>
                    <td style="padding: 40px 40px 20px 40px;">
                      <p style="color: #4a5568; font-size: 16px; margin-top: 0; line-height: 1.5;">
                        Ciao <strong>Salvatore</strong>, un nuovo utente si è appena registrato al portale:
                      </p>
                      
                      <!-- Box Nome Utente in evidenza -->
                      <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-left: 4px solid #FF64AE; padding: 15px 20px; border-radius: 6px; margin: 25px 0;">
                        <span style="font-size: 12px; color: #718096; display: block; text-transform: uppercase; letter-spacing: 1px; font-weight: 600;">Nuovo Iscritto</span>
                        <strong style="font-size: 18px; color: #1a202c; display: block; margin-top: 3px;">${nome} ${cognome}</strong>
                      </div>

                      <!-- Tabella Dettaglio Contatti -->
                      <p style="color: #4a5568; font-size: 14px; font-weight: 600; margin-bottom: 10px;">Dettagli di contatto:</p>
                      <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; overflow: hidden;">
                        <tbody>
                          <tr>
                            <td style="padding: 12px 15px; border-bottom: 1px solid #edf2f7; font-size: 14px; color: #718096; width: 30%;"><strong>Email</strong></td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid #edf2f7; font-size: 14px; color: #2d3748;">${email || "Non specificata"}</td>
                          </tr>
                          <tr>
                            <td style="padding: 12px 15px; border-bottom: 1px solid #edf2f7; font-size: 14px; color: #718096;"><strong>Telefono</strong></td>
                            <td style="padding: 12px 15px; border-bottom: 1px solid #edf2f7; font-size: 14px; color: #2d3748;">${telefono || "Non specificato"}</td>
                          </tr>
                          <tr>
                            <td style="padding: 12px 15px; font-size: 14px; color: #718096;"><strong>Registrato il</strong></td>
                            <td style="padding: 12px 15px; font-size: 14px; color: #2d3748;">${Utilities.formatDate(dataRegistrazione, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}</td>
                          </tr>
                        </tbody>
                      </table>

                      <p style="color: #718096; font-size: 14px; line-height: 1.5; margin-top: 30px;">
                        L'utente è stato aggiunto correttamente nel foglio anagrafico.
                      </p>
                    </td>
                  </tr>

                  <!-- Footer -->
                  <tr>
                    <td style="background-color: #f8fafc; padding: 20px 40px; text-align: center; border-top: 1px solid #edf2f7;">
                      <p style="color: #a0aec0; font-size: 12px; margin: 0;">
                        Generato automaticamente dalla tua Web App • Non rispondere a questa email.
                      </p>
                    </td>
                  </tr>

                </table>
              </td>
            </tr>
          </table>
        </body>
        </html>
      `;

      GmailApp.sendEmail(tuaEmail, oggetto, "", { htmlBody: corpoHtml });
    }
  } catch (err) {
    console.error("Errore nell'invio della mail di notifica registrazione: " + err.message);
  }

  return { success: true, message: "✅ Registrazione completata con successo!" };
}

function registraENuovoUtente(nome, cognome, telefono, email, rigaLezione) {
  const regRes = registraNuovoUtenteInTab(nome, cognome, telefono, email);
  if (!regRes.success && !regRes.message.includes("già registrato")) {
    return regRes;
  }
  return prenotaLezione(nome, cognome, rigaLezione);
}

function registraAcquistoUtente(telefono, categoria, numLezioni) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUtenti = ss.getSheetByName("Utenti");
  var sheetAcquisti = ss.getSheetByName("Acquisti");
  
  if (!sheetUtenti || !sheetAcquisti) {
    return { success: false, message: "Foglio Utenti o Acquisti non trovato." };
  }
  
  var telCercato = String(telefono || "").trim();
  if (!telCercato) {
    return { success: false, message: "Inserisci un numero di telefono valido." };
  }

  // 1. Controllo se l'utente esiste nel foglio Utenti (Telefono in Colonna C -> indice 2)
  var dataUtenti = sheetUtenti.getDataRange().getValues();
  var utenteTrovato = false;
  var nomeUtente = "";
  var cognomeUtente = "";
  
  for (var i = 1; i < dataUtenti.length; i++) {
    var uTel = String(dataUtenti[i][2] || "").trim();
    if (uTel === telCercato) {
      utenteTrovato = true;
      nomeUtente = dataUtenti[i][0] || "";    
      cognomeUtente = dataUtenti[i][1] || "";  
      break;
    }
  }
  
  if (!utenteTrovato) {
    return { 
      success: false, 
      code: "UTENTE_NON_REGISTRATO", 
      message: "Attenzione: questo numero di telefono non risulta registrato. Effettua prima la registrazione!" 
    };
  }
  
  // 2. Controllo nel foglio Acquisti se esiste già una riga per Telefono + Categoria
  var dataAcquisti = sheetAcquisti.getDataRange().getValues();
  var rigaEsistente = -1;
  
  for (var j = 1; j < dataAcquisti.length; j++) {
    var aTel = String(dataAcquisti[j][7] || "").trim(); 
    var aCat = String(dataAcquisti[j][3] || "").trim().toLowerCase(); 
    
    if (aTel === telCercato && aCat === String(categoria || "").trim().toLowerCase()) {
      rigaEsistente = j + 1; 
      break;
    }
  }
  
  var dataOdierna = new Date();
  var lezioniDaAggiungere = Number(numLezioni) || 0;
  var daSaldareDalFoglio = 0;
  
  if (rigaEsistente !== -1) {
    // --- AGGIORNA LA RIGA ESISTENTE ---
    var cellaAcquistate = sheetAcquisti.getRange(rigaEsistente, 5);
    var vecchieAcquistate = Number(cellaAcquistate.getValue()) || 0;
    cellaAcquistate.setValue(vecchieAcquistate + lezioniDaAggiungere);
    
    var cellaDaSaldare = sheetAcquisti.getRange(rigaEsistente, 7);
    var vecchieDaSaldare = Number(cellaDaSaldare.getValue()) || 0;
    var nuovoDaSaldare = vecchieDaSaldare + lezioniDaAggiungere;
    cellaDaSaldare.setValue(nuovoDaSaldare);
    
    sheetAcquisti.getRange(rigaEsistente, 1).setValue(dataOdierna);
    
    // Leggiamo il valore aggiornato direttamente dalla cella (Colonna G = 7) del foglio Excel
    daSaldareDalFoglio = Number(cellaDaSaldare.getValue()) || 0;
    
  } else {
    // --- CREA UNA NUOVA RIGA SE NON ESISTE ---
    sheetAcquisti.appendRow([dataOdierna, nomeUtente, cognomeUtente, categoria, lezioniDaAggiungere, 0, lezioniDaAggiungere, telCercato]);
    
    // Leggiamo il valore appena scritto nel foglio rintracciando l'ultima riga inserita
    var ultimaRiga = sheetAcquisti.getLastRow();
    daSaldareDalFoglio = Number(sheetAcquisti.getRange(ultimaRiga, 7).getValue()) || 0;
  }
  
  // --- INVIO NOTIFICA EMAIL ---
  if (typeof inviaNotificaAcquistoPacchetti === 'function') {
    const pacchettoNotifica = [{
      categoria: categoria,
      acquistate: lezioniDaAggiungere, // <-- Le lezioni appena acquistate
      daSaldare: daSaldareDalFoglio     // <-- Il totale da saldare preso direttamente dal foglio Excel
    }];
    inviaNotificaAcquistoPacchetti(nomeUtente, cognomeUtente, pacchettoNotifica);
  }
  
  return { success: true, message: "Pacchetto aggiunto con successo e importo aggiornato da saldare!" };
}



// ==========================================
// SEZIONE: LISTINI & CATEGORIE ISTRUTTORE
// ==========================================

function getConfigurazioneListini() {
  verificaInizializzazioneFogli();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Listini');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const listini = [];

  for (let i = 1; i < data.length; i++) {
    const categoria = String(data[i][0] || '').trim();
    if (!categoria) continue;

    listini.push({
      rowIndex: i + 1,
      categoria: categoria,
      p1Num: Number(data[i][1]) || 0,
      p1Prezzo: Number(data[i][2]) || 0,
      p2Num: Number(data[i][3]) || 0,
      p2Prezzo: Number(data[i][4]) || 0,
      p3Num: Number(data[i][5]) || 0,
      p3Prezzo: Number(data[i][6]) || 0
    });
  }

  return listini;
}

function salvaCategoriaListino(rowIndex, categoria, p1Num, p1Prezzo, p2Num, p2Prezzo, p3Num, p3Prezzo) {
  verificaInizializzazioneFogli();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Listini');
  
  if (rowIndex) {
    sheet.getRange(rowIndex, 1).setValue(categoria);
    sheet.getRange(rowIndex, 2).setValue(p1Num);
    sheet.getRange(rowIndex, 3).setValue(p1Prezzo);
    sheet.getRange(rowIndex, 4).setValue(p2Num);
    sheet.getRange(rowIndex, 5).setValue(p2Prezzo);
    sheet.getRange(rowIndex, 6).setValue(p3Num);
    sheet.getRange(rowIndex, 7).setValue(p3Prezzo);
    return { success: true, message: "Categoria aggiornata con successo!" };
  } else {
    sheet.appendRow([categoria, p1Num, p1Prezzo, p2Num, p2Prezzo, p3Num, p3Prezzo]);
    return { success: true, message: "Nuova categoria aggiunta con successo!" };
  }
}

function eliminaCategoriaListino(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Listini');
  if (!sheet) return { success: false };
  sheet.deleteRow(rowIndex);
  return { success: true, message: "Categoria eliminata." };
}

// ==========================================
// SEZIONE: DASHBOARD ISTRUTTORE & UTENTI
// ==========================================

function getTuttiUtenti() {
  verificaInizializzazioneFogli();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUtenti = ss.getSheetByName('Utenti');
    const sheetAcquisti = ss.getSheetByName('Acquisti'); // Oppure 'Movimenti' se lo hai chiamato così
    if (!sheetUtenti) return [];
    
    const dataUtenti = sheetUtenti.getDataRange().getValues();

    // Leggi riepilogo acquisti per utente
    const acqData = sheetAcquisti ? sheetAcquisti.getDataRange().getValues() : [];
    const mappaAcquisti = {};

    // Mappatura corretta in base alle colonne del foglio:
    // B = Nome (indice 1), C = Cognome (indice 2), D = Tipocategoria (indice 3)
    // E = Acquistate (indice 4), F = Utilizzate (indice 5), G = Da saldare (indice 6)
    for (let j = 1; j < acqData.length; j++) {
      const uNome = String(acqData[j][1] || '').trim().toLowerCase();
      const uCognome = String(acqData[j][2] || '').trim().toLowerCase();
      const chiave = `${uNome}_${uCognome}`;

      if (!mappaAcquisti[chiave]) mappaAcquisti[chiave] = [];
      mappaAcquisti[chiave].push({
        categoria: acqData[j][3],                  // Colonna D
        acquistate: Number(acqData[j][4]) || 0,    // Colonna E
        utilizzate: Number(acqData[j][5]) || 0,    // Colonna F
        daSaldare: Number(acqData[j][6]) || 0      // Colonna G
      });
    }

    const utenti = [];
    for (let i = 1; i < dataUtenti.length; i++) {
      if (!dataUtenti[i][0]) continue;
      
      const nome = String(dataUtenti[i][0] || '').trim();
      const cognome = String(dataUtenti[i][1] || '').trim();
      const telefono = String(dataUtenti[i][2] || '').trim();
      const email = String(dataUtenti[i][3] || '').trim();
      const acquistate = Number(dataUtenti[i][4]) || 0;
      const utilizzate = Number(dataUtenti[i][5]) || 0;
      const daSaldare = Number(dataUtenti[i][6]) || 0;
      
      const chiave = `${nome.toLowerCase()}_${cognome.toLowerCase()}`;

      utenti.push({
        rowIndex: i + 1,
        nome: nome,
        cognome: cognome,
        telefono: telefono,
        email: email,
        acquistate: acquistate,
        utilizzate: utilizzate,
        daSaldare: daSaldare,
        dettaglioAcquisti: mappaAcquisti[chiave] || []
      });
    }
    
    return utenti;
  } catch (err) {
    console.error("Errore in getTuttiUtenti: " + err.message);
    return [];
  }
}

function salvaUtenteIstruttore(rowIndex, nome, cognome, telefono, email, daSaldare) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Utenti');
  if (!sheet) return { success: false, message: "Tab Utenti non trovato." };

  if (rowIndex) {
    sheet.getRange(rowIndex, 1).setValue(nome);
    sheet.getRange(rowIndex, 2).setValue(cognome);
    sheet.getRange(rowIndex, 3).setValue(telefono);
    sheet.getRange(rowIndex, 4).setValue(email);
    if (daSaldare !== undefined) sheet.getRange(rowIndex, 7).setValue(daSaldare);
    return { success: true, message: "Utente aggiornato con successo!" };
  } else {
    sheet.appendRow([nome, cognome, telefono, email, 0, 0, daSaldare || 0]);
    return { success: true, message: "Nuovo utente creato con successo!" };
  }
} 

function eliminaUtenteIstruttore(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Utenti');
  if (!sheet) return { success: false, message: "Tab Utenti non trovato." };
  sheet.deleteRow(rowIndex);
  return { success: true, message: "Utente eliminato." };
}

function getIscrittiLezione(idLezione) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPren = ss.getSheetByName('Prenotazioni');
  
  if (!sheetPren) return [];

  const dataPren = sheetPren.getDataRange().getValues();
  const iscritti = [];
  const targetId = String(idLezione || "").trim();

  for (let i = 1; i < dataPren.length; i++) {
    const pNome = String(dataPren[i][0] || "").trim();
    const pCognome = String(dataPren[i][1] || "").trim();
    const pIdLezione = String(dataPren[i][8] || "").trim(); // Colonna I: ID Lezione
    const pIdPrenotazione = String(dataPren[i][9] || "").trim(); // Colonna J: ID Prenotazione

    if (pIdLezione !== "" && pIdLezione === targetId) {
      iscritti.push({
        prenotazioneRow: i + 1,
        idPrenotazione: pIdPrenotazione, // <-- Fondamentale!
        nome: pNome,
        cognome: pCognome
      });
    }
  }
  
  return iscritti;
}

function spostaPrenotazioneIstruttore(prenotazioneRow, nuovaRigaLezione) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPren = ss.getSheetByName('Prenotazioni');
  const sheetCal = ss.getSheetByName('Calendario');
  const tz = ss.getSpreadsheetTimeZone();

  const prenRowData = sheetPren.getRange(prenotazioneRow, 1, 1, 8).getValues()[0];
  const vecchioNomeLez = String(prenRowData[5]).trim();
  const vecchiaData = prenRowData[2] instanceof Date ? Utilities.formatDate(prenRowData[2], tz, 'dd/MM/yyyy') : String(prenRowData[2]);
  const vecchioOra = prenRowData[3] instanceof Date ? Utilities.formatDate(prenRowData[3], tz, 'HH:mm') : String(prenRowData[3]);

  const nuovaLezData = sheetCal.getRange(nuovaRigaLezione, 1, 1, 8).getValues()[0];
  const nuovoNomeLez = String(nuovaLezData[0]);
  const rawDate = nuovaLezData[1];
  const nuovaDataStr = rawDate instanceof Date ? Utilities.formatDate(rawDate, tz, 'dd/MM/yyyy') : String(rawDate);
  const nuovoOraStr = nuovaLezData[2] instanceof Date ? Utilities.formatDate(nuovaLezData[2], tz, 'HH:mm') : String(nuovaLezData[2]);
  const nuovoLuogo = String(nuovaLezData[4]);
  const maxPosti = Number(nuovaLezData[5]) || 0;
  let postiPren = Number(nuovaLezData[6]) || 0;
  const nuovaCategoria = String(nuovaLezData[7] || 'Generale');

  if (postiPren >= maxPosti) return { success: false, message: "La lezione di destinazione è esaurita!" };

  sheetPren.getRange(prenotazioneRow, 3).setValue(nuovaDataStr);
  sheetPren.getRange(prenotazioneRow, 4).setValue(nuovoOraStr);
  sheetPren.getRange(prenotazioneRow, 6).setValue(nuovoNomeLez);
  sheetPren.getRange(prenotazioneRow, 7).setValue(nuovoLuogo);
  sheetPren.getRange(prenotazioneRow, 8).setValue(nuovaCategoria);

  const calData = sheetCal.getDataRange().getValues();
  for (let k = 1; k < calData.length; k++) {
    const cNome = String(calData[k][0]).trim();
    const cData = calData[k][1] instanceof Date ? Utilities.formatDate(calData[k][1], tz, 'dd/MM/yyyy') : String(calData[k][1]);
    const cOra = calData[k][2] instanceof Date ? Utilities.formatDate(calData[k][2], tz, 'HH:mm') : String(calData[k][2]);
    if (cNome.toLowerCase() === vecchioNomeLez.toLowerCase() && cData === vecchiaData && cOra === vecchioOra) {
      let pCurr = Number(calData[k][6]) || 0;
      if (pCurr > 0) sheetCal.getRange(k + 1, 7).setValue(pCurr - 1);
      break;
    }
  }

  sheetCal.getRange(nuovaRigaLezione, 7).setValue(postiPren + 1);

  return { success: true, message: "Prenotazione spostata con successo!" };
}
function getCategorieListino() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listini");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var categorie = [];
  
  // Salta l'intestazione (parte da i = 1)
  for (var i = 1; i < data.length; i++) {
    var cat = data[i][0];
    if (cat && cat.toString().trim() !== "") {
      categorie.push(cat.toString().trim());
    }
  }
  return categorie;
}
function getListinoCompleto() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listini");
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var listino = {};
  
  // Supponiamo che la riga 1 contenga le intestazioni (es. Categoria, Pacchetto 1, Prezzo 1, ecc.)
  // Oppure mappiamo le colonne da B a M in base alla struttura del tuo foglio.
  // Struttura tipica: Colonna A = Categoria, poi a coppie o sequenza i pacchetti e i prezzi.
  for (var i = 1; i < data.length; i++) {
    var categoria = String(data[i][0] || "").trim();
    if (!categoria) continue;
    
    var pacchettiDisponibili = [];
    
    // Scansioniamo le colonne da B (indice 1) fino a M (indice 12)
    // Raccogliamo i valori non vuoti (puoi adattare la logica se hai nome pacchetto e prezzo alternati)
    for (var col = 1; col < data[i].length; col++) {
      var val = data[i][col];
      if (val !== "" && val !== null && val !== undefined) {
        pacchettiDisponibili.push({
          colonnaIndex: col,
          valore: val
        });
      }
    }
    
    listino[categoria] = pacchettiDisponibili;
  }
  return listino;
}
function getPacchettiPerCategoria(categoriaSelezionata) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listini");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var pacchettiTrovati = [];

  for (var i = 1; i < data.length; i++) {
    var cat = String(data[i][0] || "").trim();
    if (cat === categoriaSelezionata) {
      
      // Scorriamo le colonne a coppie partendo da B/C (indici 1 e 2), poi D/E (3 e 4), F/G (5 e 6) ecc. fino a M
      for (var j = 1; j < data[i].length - 1; j += 2) {
        var numLezioni = data[i][j];     // Colonna dispari (es. B, D, F) -> Numero lezioni
        var prezzo = data[i][j + 1];     // Colonna pari (es. C, E, G) -> Prezzo

        // Verifichiamo che ci sia un numero di lezioni valido
        if (numLezioni !== "" && numLezioni !== null && numLezioni !== undefined) {
          pacchettiTrovati.push({
            numLezioni: numLezioni,
            prezzo: prezzo !== "" ? prezzo : 0
          });
        }
      }
      break;
    }
  }
  return pacchettiTrovati;
}
function registraUtenteCompleto(dati) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUtenti = ss.getSheetByName('Utenti');
    const sheetAcquisti = ss.getSheetByName('Acquisti'); 
    
    if (!sheetUtenti) return { success: false, message: "Foglio Utenti non trovato." };
    
    const nome = String(dati.nome || "").trim();
    const cognome = String(dati.cognome || "").trim();
    const telefonoUtente = String(dati.telefono || dati.tel || "").trim();
    
    if (!nome || !cognome) return { success: false, message: "Nome e Cognome sono obbligatori." };
    
    // 1. Controllo se l'utente esiste già
    const dataUtenti = sheetUtenti.getDataRange().getValues();
    for (let i = 1; i < dataUtenti.length; i++) {
      if (String(dataUtenti[i][0]).trim().toLowerCase() === nome.toLowerCase() && 
          String(dataUtenti[i][1]).trim().toLowerCase() === cognome.toLowerCase()) {
        return { success: false, message: "Esiste già un utente registrato con questo Nome e Cognome." };
      }
    }
    
    // 2. Inserisce il nuovo utente nel foglio Utenti
    sheetUtenti.appendRow([nome, cognome, telefonoUtente, dati.email || ""]);
    
    // 3. Inserisce i pacchetti nel foglio Acquisti gestendo il telefono riga per riga come la categoria
    if (dati.pacchetti && Array.isArray(dati.pacchetti) && sheetAcquisti) {
      dati.pacchetti.forEach(p => {
        const cat = String(p.categoria || "").trim();
        if (!cat) return;
        sheetAcquisti.appendRow([
          new Date(),                // Colonna A: Data
          nome,                      // Colonna B: Nome
          cognome,                   // Colonna C: Cognome
          cat,                       // Colonna D: Categoria
          Number(p.acquistate) || 0, // Colonna E: Acquistate
          Number(p.utilizzate) || 0, // Colonna F: Utilizzate
          Number(p.daSaldare) || 0,  // Colonna G: Da saldare
          telefonoUtente                    // Colonna H: Telefono
        ]);
      });
    }
    
    // 4. Aggiorna i totali generali di riepilogo
    if (typeof aggiornaTotaliUtente === 'function') {
      aggiornaTotaliUtente(nome, cognome);
    }
    
    return { success: true, message: "Nuovo utente registrato con successo!" };
  } catch (err) {
    return { success: false, message: "Errore durante la registrazione: " + err.message };
  }
}

function inviaNotificaAcquistoPacchetti(nome, cognome, pacchetti) {
  try {
    // Legge la preferenza dalle proprietà dello script (di default è true se non impostato)
    var scriptProperties = PropertiesService.getScriptProperties();
    var inviaEmailAttivo = scriptProperties.getProperty('NOTIFICA_EMAIL_ACQUISTO');
    
    // Se l'utente ha disattivato il toggle (quindi la proprietà è "false"), blocca l'invio
    if (inviaEmailAttivo === 'false') {
      console.log("Notifiche email disattivate dall'utente. Invio saltato.");
      return;
    }

    const tuaEmail = "salvatore.vatrella@gmail.com"; 
    if (!tuaEmail) return;

    let righeTabellaHtml = '';
    let totaleDaSaldareNum = 0;

    pacchetti.forEach(p => {
      const cat = p.categoria || 'Generale';
      const acq = Number(p.acquistate) || 0;
      const saldare = Number(p.daSaldare) || 0;
      totaleDaSaldareNum += saldare;

      righeTabellaHtml += `
        <tr>
          <td style="padding: 12px 15px; border-bottom: 1px solid #edf2f7; font-size: 14px; color: #2d3748;">
            <span style="font-weight: 600; background-color: #f3e8ff; color: #FF64AE; padding: 4px 8px; border-radius: 4px; font-size: 12px;">${cat}</span>
          </td>
          <td style="padding: 12px 15px; border-bottom: 1px solid #edf2f7; font-size: 14px; color: #2d3748; text-align: center; font-weight: 500;">
            ${acq} lez.
          </td>
          <td style="padding: 12px 15px; border-bottom: 1px solid #edf2f7; font-size: 14px; color: #e53e3e; font-weight: bold; text-align: center;">
            ${saldare} da saldare
          </td>
        </tr>
      `;
    });

    const oggetto = `Nuova Ricarica Registrata: ${nome} ${cognome}`;
    
    const corpoHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
      </head>
      <body style="margin: 0; padding: 0; background-color: #f4f6f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="table-layout: fixed; background-color: #f4f6f9; padding: 30px 0;">
          <tr>
            <td align="center">
              <table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05); overflow: hidden;">
                
                <tr>
                  <td style="background: linear-gradient(135deg, #FF64AE 0%, #FF58A1 100%); padding: 35px 40px; text-align: center;">
                    <h1 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: bold; letter-spacing: 0.5px;">
                      Nuova Ricarica Registrata
                    </h1>
                    <p style="color: #f3e8ff; margin: 6px 0 0 0; font-size: 14px; font-weight: 500;">
                      Fitness con Valentina
                    </p>
                  </td>
                </tr>

                <tr>
                  <td style="padding: 40px 40px 20px 40px;">
                    <p style="color: #4a5568; font-size: 16px; margin-top: 0; line-height: 1.5;">
                      Ciao <strong>Salvatore</strong>, è stato registrato con successo un nuovo pacchetto lezioni per:
                    </p>
                    
                    <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-left: 4px solid #FF64AE; padding: 15px 20px; border-radius: 6px; margin: 25px 0;">
                      <span style="font-size: 12px; color: #718096; display: block; text-transform: uppercase; letter-spacing: 1px; font-weight: 600;">Cliente</span>
                      <strong style="font-size: 18px; color: #1a202c; display: block; margin-top: 3px;">${nome} ${cognome}</strong>
                    </div>

                    <p style="color: #4a5568; font-size: 14px; font-weight: 600; margin-bottom: 10px;">Dettaglio transazione:</p>
                    <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; overflow: hidden;">
                      <thead>
                        <tr style="background-color: #f8fafc;">
                          <th style="padding: 12px 15px; text-align: left; font-size: 12px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e2e8f0;">Categoria</th>
                          <th style="padding: 12px 15px; text-align: center; font-size: 12px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e2e8f0;">Acquistate</th>
                          <th style="padding: 12px 15px; text-align: center; font-size: 12px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e2e8f0;">Stato</th>
                        </tr>
                      </thead>
                      <tbody>
                        ${righeTabellaHtml}
                      </tbody>
                    </table>

                    <p style="color: #718096; font-size: 14px; line-height: 1.5; margin-top: 30px;">
                      Ricordati di riscuotere il pagamento alla prossima lezione.
                    </p>
                  </td>
                </tr>

                <tr>
                  <td style="background-color: #f8fafc; padding: 20px 40px; text-align: center; border-top: 1px solid #edf2f7;">
                    <p style="color: #a0aec0; font-size: 12px; margin: 0;">
                      Generato automaticamente dalla tua Web App • Non rispondere a questa email.
                    </p>
                  </td>
                </tr>

              </table>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    GmailApp.sendEmail(tuaEmail, oggetto, "", { htmlBody: corpoHtml });
  } catch (err) {
    console.error("Errore nell'invio della mail di notifica: " + err.message);
  }
}

// Funzione chiamata da JavaScript quando sposti il toggle
function salvaStatoNotificheEmail(attivo) {
  PropertiesService.getScriptProperties().setProperty('NOTIFICA_EMAIL_ACQUISTO', attivo);
  return "Salvato con successo";
}

// Funzione chiamata all'avvio della web app per impostare il toggle nella posizione corretta
function getStatoNotificheEmail() {
  var val = PropertiesService.getScriptProperties().getProperty('NOTIFICA_EMAIL_ACQUISTO');
  // Se non è mai stato salvato, restituisce true di default
  return val === null ? true : (val === 'true');
}
function testInvioMail() {
  GmailApp.sendEmail(
    "salvatore.vatrella@gmail.com", 
    "Test Autorizzazione Gmail", 
    "Questa è una mail di prova per autorizzare gli script di Google a inviare email."
  );
  Logger.log("Email di prova inviata con successo!");
}

// Restituisce gli utenti che hanno almeno 1 credito residuo per la categoria richiesta
function getUtentiConCreditiPerCategoria(categoriaLezione) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAcquisti = ss.getSheetByName("Acquisti");
  if (!sheetAcquisti) return [];

  var data = sheetAcquisti.getDataRange().getValues();
  var utentiDisponibili = [];
  var catC = String(categoriaLezione || "").trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var nome = String(data[i][1] || "").trim();
    var cognome = String(data[i][2] || "").trim();
    var cat = String(data[i][3] || "").trim().toLowerCase();
    var acquistate = Number(data[i][4]) || 0;
    var utilizzate = Number(data[i][5]) || 0;
    var rimanenti = acquistate - utilizzate;

    // Se la categoria corrisponde (o se vuoi gestirla in modo flessibile) e ci sono crediti
    if (cat === catC && rimanenti > 0) {
      // Evitiamo duplicati se un utente ha più righe per la stessa categoria
      var esisteGia = utentiDisponibili.some(u => u.nome.toLowerCase() === nome.toLowerCase() && u.cognome.toLowerCase() === cognome.toLowerCase());
      if (!esisteGia) {
        utentiDisponibili.push({
          nome: nome,
          cognome: cognome,
          categoria: cat,
          rimanenti: rimanenti
        });
      }
    }
  }
  return utentiDisponibili;
}

// Iscrive l'utente selezionato e scala 1 credito
function amministratoreIscriviUtente(nome, cognome, categoriaLezione, dataLezione, oraInizio, oraFine, lezioneNome, luogo, lezioneId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAcquisti = ss.getSheetByName("Acquisti");
  var sheetPrenotazioni = ss.getSheetByName("Prenotazioni");
  var sheetCalendario = ss.getSheetByName("Calendario") || ss.getSheetByName("Lezioni"); 
  var sheetUtenti = ss.getSheetByName("Utenti") || ss.getSheetByName("Clienti"); // Per recuperare il telefono se presente

  if (!sheetAcquisti || !sheetPrenotazioni) {
    return { success: false, message: "Fogli di sistema non trovati." };
  }

  var dataFormattata = "";
  var dStr = String(dataLezione || "").trim();

  if (dStr.indexOf("T") !== -1) {
    dStr = dStr.split("T")[0];
  } else if (dStr.indexOf(" ") !== -1) {
    dStr = dStr.split(" ")[0];
  }

  // Conversione sicura nel formato dd/mm/yyyy
  if (dStr.indexOf("-") !== -1) {
    var parti = dStr.split("-");
    if (parti.length === 3) {
      dataFormattata = parti[2] + "/" + parti[1] + "/" + parti[0]; 
    }
  } else if (dStr.indexOf("/") !== -1) {
    dataFormattata = dStr;
  }

  if (!dataFormattata) {
    dataFormattata = Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy");
  }

  // 1. Cerca la lezione nel foglio Calendario
  var rigaCalendario = -1;
  var postiMax = 0;
  var postiPrenotati = 0;

  if (sheetCalendario) {
    var calData = sheetCalendario.getDataRange().getValues();
    var cNome = String(lezioneNome || "").trim().toLowerCase();
    var cOraInizio = String(oraInizio || "").trim();
    var cLuogo = String(luogo || "").trim().toLowerCase();
    var cDataTarget = dataFormattata; 

    for (var j = 1; j < calData.length; j++) {
      var rowNome = String(calData[j][0] || "").trim().toLowerCase();
      var rawRowDate = calData[j][1];
      var rowDataStr = "";

      if (rawRowDate instanceof Date) {
        rowDataStr = Utilities.formatDate(rawRowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        rowDataStr = String(rawRowDate || "").trim().substring(0, 10);
      }

      var rowOraInizio = String(calData[j][2] || "").trim();
      var rowLuogo = String(calData[j][4] || "").trim().toLowerCase();
      
      // Controllo prioritario sull'ID Lezione se passato, altrimenti corrispondenza dettagliata
      var rowIdLezione = String(calData[j][8] || "").trim();
      var matchId = lezioneId && rowIdLezione === String(lezioneId).trim();
      var matchDettagli = (rowNome === cNome && rowDataStr === cDataTarget && rowOraInizio === cOraInizio && rowLuogo === cLuogo);

      if (matchId || matchDettagli) {
        rigaCalendario = j + 1;                             
        postiMax = Number(calData[j][5]) || 0;       
        postiPrenotati = Number(calData[j][6]) || 0; 
        
        // Se l'idLezione mancava in input ma è presente nel foglio, lo recuperiamo
        if (!lezioneId && rowIdLezione) {
          lezioneId = rowIdLezione;
        }
        break;
      }
    }

    if (rigaCalendario !== -1) {
      if (postiMax > 0 && postiPrenotati >= postiMax) {
        return { success: false, message: "Impossibile iscrivere l'utente: la lezione è al completo." };
      }
    }
  }

  // 2. Controllo crediti nel foglio Acquisti e recupero telefono
  var nomeC = String(nome || "").trim().toLowerCase();
  var cognomeC = String(cognome || "").trim().toLowerCase();
  var catC = String(categoriaLezione || "").trim().toLowerCase();

  var acqData = sheetAcquisti.getDataRange().getValues();
  var rigaAcquisto = -1;
  var rimanenti = 0;
  var telefonoUtente = "";

  for (var i = 1; i < acqData.length; i++) {
    var aNome = String(acqData[i][1] || "").trim().toLowerCase();
    var aCognome = String(acqData[i][2] || "").trim().toLowerCase();
    var aCat = String(acqData[i][3] || "").trim().toLowerCase();
    var aTel = String(acqData[i][7] || "").trim(); // Supponendo colonna H (indice 7) per il telefono in Acquisti

    if (aTel && !telefonoUtente && aNome === nomeC && aCognome === cognomeC) {
      telefonoUtente = aTel;
    }

    if (aNome === nomeC && aCognome === cognomeC && aCat === catC) {
      var acquistate = Number(acqData[i][4]) || 0;
      var utilizzate = Number(acqData[i][5]) || 0;
      rimanenti = acquistate - utilizzate;

      if (rimanenti > 0) {
        rigaAcquisto = i + 1;
        break;
      }
    }
  }

  // Se non trovato in Acquisti, proviamo a cercare il telefono nel foglio Utenti/Clienti
  if (!telefonoUtente && sheetUtenti) {
    var utentiData = sheetUtenti.getDataRange().getValues();
    for (var u = 1; u < utentiData.length; u++) {
      var uNome = String(utentiData[u][0] || "").trim().toLowerCase();
      var uCognome = String(utentiData[u][1] || "").trim().toLowerCase();
      var uTel = String(utentiData[u][2] || "").trim(); // Regola in base al tuo indice colonna telefono
      if (uNome === nomeC && uCognome === cognomeC && uTel) {
        telefonoUtente = uTel;
        break;
      }
    }
  }

  if (rigaAcquisto === -1) {
    return { success: false, message: "L'utente non ha crediti disponibili per questa categoria." };
  }

  // 3. Scala 1 credito nel foglio Acquisti
  var cellaUtilizzate = sheetAcquisti.getRange(rigaAcquisto, 6);
  var vecchieUtilizzate = Number(cellaUtilizzate.getValue()) || 0;
  cellaUtilizzate.setValue(vecchieUtilizzate + 1);

  // 4. Generazione ID Prenotazione Univoco
  var idPrenotazioneUnivoco = Utilities.getUuid();

  // 5. Scrittura nel foglio Prenotazioni con campi ID Lezione, ID Prenotazione e Telefono
  var ultimaRiga = sheetPrenotazioni.getLastRow() + 1;

  sheetPrenotazioni.getRange(ultimaRiga, 1).setValue(nome);
  sheetPrenotazioni.getRange(ultimaRiga, 2).setValue(cognome);

  var cellaData = sheetPrenotazioni.getRange(ultimaRiga, 3);
  cellaData.setNumberFormat("@"); 
  cellaData.setValue(dataFormattata + ""); 

  sheetPrenotazioni.getRange(ultimaRiga, 4).setValue(oraInizio || "");
  sheetPrenotazioni.getRange(ultimaRiga, 5).setValue(oraFine || "");
  sheetPrenotazioni.getRange(ultimaRiga, 6).setValue(lezioneNome || "Lezione");
  sheetPrenotazioni.getRange(ultimaRiga, 7).setValue(luogo || "");
  sheetPrenotazioni.getRange(ultimaRiga, 8).setValue(categoriaLezione);
  
  // Colonna I (9): ID Lezione
  sheetPrenotazioni.getRange(ultimaRiga, 9).setValue(lezioneId || "");
  
  // Colonna J (10): ID Prenotazione (Univoco)
  sheetPrenotazioni.getRange(ultimaRiga, 10).setValue(idPrenotazioneUnivoco);
  
  // Colonna K (11): Telefono
  sheetPrenotazioni.getRange(ultimaRiga, 11).setValue(telefonoUtente);

  // 6. Incrementa i posti prenotati nel foglio Calendario
  if (rigaCalendario !== -1) {
    var cellaPostiPrenotati = sheetCalendario.getRange(rigaCalendario, 7); 
    cellaPostiPrenotati.setValue(postiPrenotati + 1);
  }

  return { success: true, message: `Utente ${nome} ${cognome} iscritto con successo!` };
}


function getListiniMappa() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Listini');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const mappa = {};
  for (let i = 1; i < data.length; i++) {
    const tipologia = String(data[i][0]).trim();
    if (tipologia) {
      mappa[tipologia] = {
        num1: data[i][1], prezzo1: data[i][2],
        num2: data[i][3], prezzo2: data[i][4],
        num3: data[i][5], prezzo3: data[i][6]
      };
    }
  }
  return mappa;
}

// Restituisce la password attuale (di default '1234' se non è mai stata cambiata)
function getPasswordIstruttore() {
  var pass = PropertiesService.getScriptProperties().getProperty('INSTRUCTOR_PASSWORD');
  return pass === null ? '1234' : pass;
}

// Verifica la password inserita al login
function verificaPasswordServer(passwordInserita) {
  var passwordCorretta = getPasswordIstruttore();
  return passwordInserita === passwordCorretta;
}

// Cambia la password verificando la vecchia
function cambiaPasswordIstruttore(vecchiaPass, nuovaPass) {
  var passwordCorretta = getPasswordIstruttore();
  
  if (vecchiaPass !== passwordCorretta) {
    return { success: false, message: "La vecchia password non è corretta." };
  }
  
  if (!nuovaPass || nuovaPass.trim().length < 4) {
    return { success: false, message: "La nuova password deve avere almeno 4 caratteri." };
  }
  
  PropertiesService.getScriptProperties().setProperty('INSTRUCTOR_PASSWORD', nuovaPass.trim());
  return { success: true, message: "Password aggiornata con successo!" };
}