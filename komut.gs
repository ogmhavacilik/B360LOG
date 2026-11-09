/**
 * GOOGLE APPS SCRIPT KODU (Code.gs) - VERSİYON 3.4 (TEKNİSYEN SOYADI İLE YAZMA)
 * -------------------------------------------------------------
 * Bu kodu Google E-Tablonuzdaki:
 * "Uzantılar > Apps Script" (Extensions > Apps Script) bölümüne
 * girip Code.gs içeriğini SİLİP bu kodun TAMAMINI yapıştırın:
 * 
 * 1. "Kaydet" (Ctrl+S) simgesine basınız.
 * 2. Sağ üstteki mavi "Dağıt" (Deploy) > "Dağıtımları Yönet" (Manage Deployments)
 * 3. Kalem (Düzenle) ikonuna tıklayın.
 * 4. Sürüm (Version) kısmından "Yeni Sürüm" (New Version) seçin.
 * 5. "Dağıt" (Deploy) butonuna basarak tamamlayın.
 * 
 * YAPILAN KRİTİK İYİLEŞTİRMELER:
 * 1. [TEKNİSYEN / OPERATÖR SOYADLARI YAZILIR]: E-Tabloya kayıt yazılırken veya güncellenirken
 *    teknisyen ve operatör isimleri sadece soyadları olarak yazılır (örn: FERHAT ÖZCAN -> ÖZCAN).
 * 2. [VERİ DOĞRULAMA HATASI ÇÖZÜLDÜ]: Google E-Tablo'daki açılır menü kısıtlamaları (clearDataValidations)
 *    güncellenen satır için kaldırılır, "C312 hücresine girdiğiniz veriler kuralları ihlal ediyor" hatası oluşmaz.
 * 3. [SATIR YERİNDE GÜNCELLEME]: Seçilen Sıra No'lu satır kendi yerinde güncellenir, yeni satır açılmaz.
 * 4. [TAM SİLME]: Silme işleminde seçilen satır tablodan tamamen kaldırılır.
 */

function getTargetSheet(ss) {
  var ssObj = ss || SpreadsheetApp.getActiveSpreadsheet();
  var active = ssObj ? ssObj.getActiveSheet() : null;
  if (active && active.getLastRow() > 1) {
    return active;
  }
  var sheets = ssObj.getSheets();
  var best = sheets[0];
  var maxR = 0;
  for (var i = 0; i < sheets.length; i++) {
    var r = sheets[i].getLastRow();
    if (r > maxR) {
      maxR = r;
      best = sheets[i];
    }
  }
  return best;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000); // 15 saniye kilit koruması

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getTargetSheet(ss);
    
    var rawData = e.postData ? e.postData.contents : "";
    var body = {};

    try {
      body = JSON.parse(rawData);
    } catch (parseErr) {
      body = e.parameter || {};
    }

    var action = (body.action || "").toString().toLowerCase().trim();
    var rawId = body.id || (body.data && body.data.id) || (body.ids && body.ids[0]) || "";
    var targetId = rawId.toString().trim();
    var targetNum = parseInt(targetId, 10);

    // E-Tablodaki tüm satırları tara (Sıra No = Sütun A)
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // Yardımcı: Sıra No eşleşen satırı bul (1-tabanlı)
    function findRowIndex() {
      if (!targetId && isNaN(targetNum)) return -1;
      for (var r = 1; r < values.length; r++) { // Başlık satırını atla (r=0)
        var cellVal = values[r][0];
        if (cellVal === undefined || cellVal === null || cellVal === "") continue;
        
        var cellStr = cellVal.toString().trim();
        var cellNum = parseInt(cellStr, 10);

        if (cellStr === targetId || (!isNaN(targetNum) && !isNaN(cellNum) && cellNum === targetNum)) {
          return r + 1; // 1-tabanlı getRange indeksi
        }
      }
      return -1;
    }

    var item = body.data || body || {};

    // Akıllı Değer Alma Yardımcısı (Personel Adları ve Tüm Alanlar İçin Esnek Arama)
    function getVal(primaryKey, altKeys) {
      if (item && item[primaryKey] !== undefined && item[primaryKey] !== null && item[primaryKey] !== "") return item[primaryKey];
      if (body && body[primaryKey] !== undefined && body[primaryKey] !== null && body[primaryKey] !== "") return body[primaryKey];
      
      if (altKeys && Array.isArray(altKeys)) {
        for (var i = 0; i < altKeys.length; i++) {
          var k = altKeys[i];
          if (item && item[k] !== undefined && item[k] !== null && item[k] !== "") return item[k];
          if (body && body[k] !== undefined && body[k] !== null && body[k] !== "") return body[k];
        }
      }

      var allKeys = [];
      if (item) for (var prop in item) allKeys.push(prop);
      if (body) for (var prop in body) if (allKeys.indexOf(prop) === -1) allKeys.push(prop);

      var normTarget = primaryKey.toLowerCase().replace(/[^a-z0-9]/g, '');
      for (var i = 0; i < allKeys.length; i++) {
        var prop = allKeys[i];
        var normProp = prop.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (normProp === normTarget) {
          var val = item[prop] !== undefined ? item[prop] : body[prop];
          if (val !== undefined && val !== null && val !== "") return val;
        }
      }
      return "";
    }

    // Soyadı Alma Yardımcısı (Sadece Teknisyen ve Operatörler için)
    function getSurname(fullName) {
      if (!fullName || typeof fullName !== "string") return "";
      var trimmed = fullName.trim();
      if (!trimmed) return "";
      var parts = trimmed.split(/\s+/);
      return parts[parts.length - 1].toUpperCase();
    }

    // ==========================================
    // 1. İŞLEM: DÜZENLEME / GÜNCELLEME (UPDATE)
    // ==========================================
    if (action === "update") {
      if (!targetId && isNaN(targetNum)) {
        return createJsonResponse({ 
          status: "error", 
          action: "update",
          message: "Güncellenecek Sıra No (id) belirtilmedi." 
        });
      }

      var rowToUpdate = findRowIndex();

      if (rowToUpdate <= 0) {
        return createJsonResponse({ 
          status: "error", 
          action: "update", 
          id: targetId, 
          message: "HATA: Sıra No " + targetId + " e-tabloda bulunamadı! Yeni satır açılmadı." 
        });
      }

      var updatedRowValues = [
        targetId,
        getVal('tarih', ['Tarih']),
        getVal('kaptanPilot', ['Kaptan Pilot', 'kaptan', 'pilot1']),
        getVal('ikinciPilot', ['2. Pilot', 'ikinciPilot', 'pilot2']),
        getSurname(getVal('teknisyen1', ['Teknisyen 1', 'teknisyen', 'teknisyen1'])),
        getSurname(getVal('operator1', ['Operator 1', 'operatör', 'operator', 'operator1'])),
        getSurname(getVal('teknisyen2', ['Teknisyen 2', 'teknisyen2'])),
        getSurname(getVal('operator2', ['Operator 2', 'operator2', 'operatör2'])),
        getVal('gorevTipi', ['Görev Tipi', 'gorevTipi']),
        getVal('gorevBolgesi', ['Görev Bölgesi', 'gorevBolgesi']),
        getVal('kalkis', ['Kalkış', 'kalkis']),
        getVal('inis', ['İniş', 'inis']),
        getVal('ucusSuresi', ['Uçuş Süresi', 'süre']),
        getVal('k9YanginHektar', ['TK-9 Yangın(Hektar)', 'yangin']),
        getVal('miktarCekim', ['Miktar (Çekim)', 'cekim']),
        getVal('tk9GorevHektar', ['TK-9 gorev(Hektar)', 'gorevHk']),
        getVal('uyduDk', ['Uydu (Dk)', 'uydu']),
        getVal('aciklama', ['AÇIKLAMA', 'aciklama'])
      ];

      var targetRange = sheet.getRange(rowToUpdate, 1, 1, updatedRowValues.length);

      // KRİTİK: Açılır menü / Veri Doğrulama (Data Validation) kurallarını bu satır için kaldır
      try {
        targetRange.clearDataValidations();
      } catch (valErr) {}

      // Verileri doğrudan aynı satır üzerine yaz
      try {
        targetRange.setValues([updatedRowValues]);
      } catch (setErr) {
        // Hücre hücre güvenli fallback yazımı
        for (var c = 0; c < updatedRowValues.length; c++) {
          try {
            var cell = sheet.getRange(rowToUpdate, c + 1);
            try { cell.clearDataValidations(); } catch (cvErr) {}
            cell.setValue(updatedRowValues[c]);
          } catch (cErr) {}
        }
      }

      SpreadsheetApp.flush();

      return createJsonResponse({ 
        status: "success", 
        action: "update", 
        id: targetId, 
        row: rowToUpdate,
        message: "Sıra No " + targetId + " olan uçuş kaydı E-Tabloda (Satır " + rowToUpdate + ") başarıyla güncellendi." 
      });
    }

    // ==========================================
    // 2. İŞLEM: SİLME (DELETE)
    // ==========================================
    if (action === "delete") {
      if (!targetId && isNaN(targetNum)) {
        return createJsonResponse({ 
          status: "error", 
          action: "delete", 
          message: "Silinecek Sıra No (id) belirtilmedi." 
        });
      }

      var rowToDelete = findRowIndex();

      if (rowToDelete > 0) {
        sheet.deleteRow(rowToDelete);
        SpreadsheetApp.flush();
        return createJsonResponse({ 
          status: "success", 
          action: "delete", 
          id: targetId, 
          message: "Sıra No " + targetId + " olan kayıt e-tablodan başarıyla silindi." 
        });
      } else {
        return createJsonResponse({ 
          status: "error", 
          action: "delete", 
          id: targetId, 
          message: "Sıra No " + targetId + " bulunamadı (zaten silinmiş olabilir)." 
        });
      }
    }

    // ==========================================
    // 3. İŞLEM: YENİ KAYIT EKLEME (ADD)
    // ==========================================
    if (action === "add" || !action) {
      var addData = body.data || body;
      
      var maxId = 0;
      for (var m = 1; m < values.length; m++) {
        var valNum = parseInt(values[m][0], 10);
        if (!isNaN(valNum) && valNum > maxId) {
          maxId = valNum;
        }
      }
      var nextId = (getVal('id') || (maxId > 0 ? maxId + 1 : values.length)).toString();

      var newRow = [
        nextId,
        getVal('tarih', ['Tarih']),
        getVal('kaptanPilot', ['Kaptan Pilot', 'kaptan']),
        getVal('ikinciPilot', ['2. Pilot', 'ikinciPilot']),
        getSurname(getVal('teknisyen1', ['Teknisyen 1', 'teknisyen'])),
        getSurname(getVal('operator1', ['Operator 1', 'operatör', 'operator'])),
        getSurname(getVal('teknisyen2', ['Teknisyen 2'])),
        getSurname(getVal('operator2', ['Operator 2'])),
        getVal('gorevTipi', ['Görev Tipi']),
        getVal('gorevBolgesi', ['Görev Bölgesi']),
        getVal('kalkis', ['Kalkış']),
        getVal('inis', ['İniş']),
        getVal('ucusSuresi', ['Uçuş Süresi']),
        getVal('k9YanginHektar', ['TK-9 Yangın(Hektar)']),
        getVal('miktarCekim', ['Miktar (Çekim)']),
        getVal('tk9GorevHektar', ['TK-9 gorev(Hektar)']),
        getVal('uyduDk', ['Uydu (Dk)']),
        getVal('aciklama', ['AÇIKLAMA'])
      ];

      try {
        sheet.appendRow(newRow);
        var lr = sheet.getLastRow();
        try { sheet.getRange(lr, 1, 1, newRow.length).clearDataValidations(); } catch(e) {}
      } catch(appErr) {
        var lr2 = sheet.getLastRow() + 1;
        var rRange = sheet.getRange(lr2, 1, 1, newRow.length);
        try { rRange.clearDataValidations(); } catch(e) {}
        rRange.setValues([newRow]);
      }

      SpreadsheetApp.flush();

      return createJsonResponse({ 
        status: "success", 
        action: "add", 
        id: nextId, 
        message: "Yeni kayıt başarıyla eklendi." 
      });
    }

    return createJsonResponse({ status: "error", message: "Bilinmeyen işlem türü: " + action });

  } catch (err) {
    return createJsonResponse({ 
      status: "error", 
      message: err.toString() 
    });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  // 1. Excel İndirme İsteği (Base64)
  if (action === "download") {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var id = ss.getId();
      var url = "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx";
      var response = UrlFetchApp.fetch(url, {
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      var blob = response.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      return createJsonResponse({ status: "success", base64: base64 });
    } catch (err) {
      return createJsonResponse({ status: "error", message: err.toString() });
    }
  }

  // 2. Varsayılan: JSON Veri Listesi
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTargetSheet(ss);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return createJsonResponse([]);
  }

  var resultList = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var siraNo = row[0];
    if (siraNo === "" || siraNo === undefined || siraNo === null) continue;

    resultList.push({
      siraNo: siraNo,
      "Sıra No": siraNo,
      tarih: row[1],
      "Tarih": row[1],
      kaptanPilot: row[2],
      "Kaptan Pilot": row[2],
      ikinciPilot: row[3],
      "2. Pilot": row[3],
      teknisyen1: row[4],
      "Teknisyen 1": row[4],
      operator1: row[5],
      "Operator 1": row[5],
      teknisyen2: row[6],
      "Teknisyen 2": row[6],
      operator2: row[7],
      "Operator 2": row[7],
      gorevTipi: row[8],
      "Görev Tipi": row[8],
      gorevBolgesi: row[9],
      "Görev Bölgesi": row[9],
      kalkis: row[10],
      "Kalkış": row[10],
      inis: row[11],
      "İniş": row[11],
      ucusSuresi: row[12],
      "Uçuş Süresi": row[12],
      k9YanginHektar: row[13],
      "TK-9 Yangın(Hektar)": row[13],
      miktarCekim: row[14],
      "Miktar (Çekim)": row[14],
      tk9GorevHektar: row[15],
      "TK-9 gorev(Hektar)": row[15],
      uyduDk: row[16],
      "Uydu (Dk)": row[16],
      aciklama: row[17],
      "AÇIKLAMA": row[17]
    });
  }

  return createJsonResponse(resultList);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
