/**
 * Google Apps Script for OGM B360 Flight Log System
 * Target Spreadsheet: https://docs.google.com/spreadsheets/d/1LqEXuXIOb0mfBRzfXJQ2skLQNT82k2uns98b-rESvkk/edit
 */

const SPREADSHEET_ID = '1LqEXuXIOb0mfBRzfXJQ2skLQNT82k2uns98b-rESvkk';

function doGet(e) {
  // DriveApp reference to ensure the script request the correct scopes
  try { DriveApp.getFileById(SPREADSHEET_ID); } catch(e) {}
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // EĞER İNDİRME İSTEĞİ GELİRSE (DİREKT EXCEL OLARAK)
  if (e && e.parameter && e.parameter.action === 'download') {
    try {
      // Belirli bir sayfa (gid) isteniyorsa onu ekle, yoksa tüm kitabı indirir
      const gid = e.parameter.gid || '1476570479'; 
      const url = "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/export?format=xlsx&gid=" + gid;
      const token = ScriptApp.getOAuthToken();
      
      const response = UrlFetchApp.fetch(url, {
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        throw new Error("Excel dışa aktarma hatası (Kod: " + response.getResponseCode() + "): " + response.getContentText());
      }
      
      const blob = response.getBlob();
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        fileName: "B360_UCUS_TABLOSU.xlsx",
        base64: Utilities.base64Encode(blob.getBytes())
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // NORMAL VERİ ÇEKME (Dashboard için)
  try {
    const sheets = ss.getSheets();
    let sheet = sheets[0];
    for (let s of sheets) {
      if (s.getRange(1, 1).getValue().toString().includes('Sıra')) {
        sheet = s;
        break;
      }
    }
    
    const data = sheet.getDataRange().getValues();
    const rawHeaders = data[0];
    
    const normalizeHeader = (h) => {
      if (!h) return '';
      return h.toString().toLowerCase()
        .replace(/ı/g, 'i').replace(/ğ/g, 'g').replace(/ü/g, 'u').replace(/ş/g, 's').replace(/ö/g, 'o').replace(/ç/g, 'c')
        .replace(/\s+(.)/g, (match, group1) => group1.toUpperCase())
        .replace(/[^\w]/g, '');
    };
    
    const uniqueHeaders = [];
    const rawUniqueHeaders = [];
    const headerCounts = {};
    const rawHeaderCounts = {};

    for (let index = 0; index < rawHeaders.length; index++) {
      let rawH = rawHeaders[index] ? rawHeaders[index].toString().trim() : '';
      let norm = normalizeHeader(rawH);
      if (!norm) {
        norm = 'col_' + index;
      }
      
      if (headerCounts[norm] !== undefined) {
        headerCounts[norm]++;
        uniqueHeaders.push(norm + headerCounts[norm]);
      } else {
        headerCounts[norm] = 1;
        uniqueHeaders.push(norm);
      }

      if (rawH) {
        if (rawHeaderCounts[rawH] !== undefined) {
          rawHeaderCounts[rawH]++;
          rawUniqueHeaders.push(rawH + ' ' + rawHeaderCounts[rawH]);
        } else {
          rawHeaderCounts[rawH] = 1;
          rawUniqueHeaders.push(rawH);
        }
      } else {
        rawUniqueHeaders.push('Col ' + index);
      }
    }
    
    const jsonData = [];
    
    for (let i = 1; i < data.length; i++) {
      let obj = {};
      let hasData = false;
      data[i].forEach((val, index) => {
        if (val !== "" && val !== null) hasData = true;
        obj['col_' + index] = val;
        let key = uniqueHeaders[index];
        let rawKey = rawUniqueHeaders[index];
        obj[key] = val;
        obj[rawKey] = val;
      });
      if (hasData) jsonData.push(obj);
    }
    
    return ContentService.createTextOutput(JSON.stringify(jsonData))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheets()[0];
    const sheets = ss.getSheets();
    for (let s of sheets) {
      if (s.getRange(1, 1).getValue().toString().includes('Sıra')) {
        sheet = s;
        break;
      }
    }
    
    let params;
    try {
      params = JSON.parse(e.postData.contents);
    } catch (f) {
      params = e.parameter;
    }

    // 1. DELETE ACTION
    if (params && params.action === 'delete') {
      const idsToDelete = [].concat(params.ids).map(function(id) { 
        return id ? id.toString().trim() : ''; 
      });
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      let deletedCount = 0;
      
      // Go backwards from bottom to top to preserve index mappings
      for (let i = values.length - 1; i >= 1; i--) {
        const rowId = values[i][0] ? values[i][0].toString().trim() : '';
        if (idsToDelete.indexOf(rowId) !== -1) {
          sheet.deleteRow(i + 1); // sheet is 1-indexed
          deletedCount++;
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: deletedCount + ' kayıt başarıyla silindi' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. UPDATE ACTION
    if (params && params.action === 'update') {
      const targetId = params.id ? params.id.toString().trim() : '';
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      let foundIndex = -1;
      
      for (let i = 1; i < values.length; i++) {
        const rowId = values[i][0] ? values[i][0].toString().trim() : '';
        if (rowId === targetId) {
          foundIndex = i + 1; // sheet is 1-indexed
          break;
        }
      }
      
      if (foundIndex === -1) {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: 'error', 
          message: 'Güncellenecek kayıt bulunamadı: ID ' + targetId 
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const updateData = params.data || {};
      
      sheet.getRange(foundIndex, 1).setValue(targetId);
      sheet.getRange(foundIndex, 2).setValue(updateData.tarih !== undefined ? updateData.tarih : '');
      sheet.getRange(foundIndex, 3).setValue(updateData.kaptanPilot !== undefined ? updateData.kaptanPilot : '');
      sheet.getRange(foundIndex, 4).setValue(updateData.ikinciPilot !== undefined ? updateData.ikinciPilot : '');
      sheet.getRange(foundIndex, 5).setValue(updateData.teknisyen1 !== undefined ? updateData.teknisyen1 : '');
      sheet.getRange(foundIndex, 6).setValue(updateData.operator1 !== undefined ? updateData.operator1 : '');
      sheet.getRange(foundIndex, 7).setValue(updateData.teknisyen2 !== undefined ? updateData.teknisyen2 : '');
      sheet.getRange(foundIndex, 8).setValue(updateData.operator2 !== undefined ? updateData.operator2 : '');
      sheet.getRange(foundIndex, 9).setValue(updateData.gorevTipi !== undefined ? updateData.gorevTipi : '');
      sheet.getRange(foundIndex, 10).setValue(updateData.gorevBolgesi !== undefined ? updateData.gorevBolgesi : '');
      sheet.getRange(foundIndex, 11).setValue(updateData.kalkis !== undefined ? updateData.kalkis : '');
      sheet.getRange(foundIndex, 12).setValue(updateData.inis !== undefined ? updateData.inis : '');
      sheet.getRange(foundIndex, 13).setValue(updateData.ucusSuresi !== undefined ? updateData.ucusSuresi : '');
      sheet.getRange(foundIndex, 14).setValue(updateData.k9YanginHektar !== undefined ? Number(updateData.k9YanginHektar) : 0);
      sheet.getRange(foundIndex, 15).setValue(updateData.miktarCekim !== undefined ? Number(updateData.miktarCekim) : 0);
      sheet.getRange(foundIndex, 16).setValue(updateData.tk9GorevHektar !== undefined ? Number(updateData.tk9GorevHektar) : 0);
      sheet.getRange(foundIndex, 17).setValue(updateData.uyduDk !== undefined ? Number(updateData.uyduDk) : 0);
      sheet.getRange(foundIndex, 18).setValue(updateData.aciklama !== undefined ? updateData.aciklama : '');
      
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Kayıt başarıyla güncellendi' 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 3. CREATE ACTION (DEFAULT APPEND)
    const lastRow = sheet.getLastRow();
    let nextId = 281; 
    if (lastRow > 1) {
      const lastIdValue = sheet.getRange(lastRow, 1).getValue();
      if (!isNaN(parseInt(lastIdValue))) {
        nextId = parseInt(lastIdValue) + 1;
      }
    }
    
    const rowData = [
      nextId,                     // A: Sıra No
      params.tarih || '',         // B: Tarih
      params.kaptanPilot || '',   // C: Kaptan Pilot
      params.ikinciPilot || '',   // D: 2. Pilot
      params.teknisyen1 || '',    // E: Teknisyen 1
      params.operator1 || '',     // F: Operator 1
      params.teknisyen2 || '',    // G: Teknisyen 2
      params.operator2 || '',     // H: Operator 2
      params.gorevTipi || '',     // I: Görev Tipi
      params.gorevBolgesi || '',  // J: Görev Bölgesi
      params.kalkis || '',        // K: Kalkış
      params.inis || '',          // L: İniş
      params.ucusSuresi || '',    // M: Uçuş Süresi
      params.k9YanginHektar || 0, // N: K9 Yangın
      params.miktarCekim || 0,    // O: Miktar Çekim
      params.tk9GorevHektar || 0, // P: TK9 Görev
      params.uyduDk || 0,         // Q: Uydu dk
      params.aciklama || ''       // R: Açıklama
    ];
    
    sheet.appendRow(rowData);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'success', 
      id: nextId, 
      message: 'Kayıt başarıyla eklendi' 
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
