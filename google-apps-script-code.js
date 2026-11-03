/**
 * GOOGLE APPS SCRIPT KODU (Code.gs)
 * -------------------------------------------------------------
 * Bu kodu Google E-Tablonuzdaki "Uzantılar > Apps Script" (Extensions > Apps Script)
 * bölümündeki Code.gs dosyasına tamamen yapıştırıp:
 * 1. "Kaydet" (Ctrl+S) butonuna basınız.
 * 2. Sağ üstteki "Dağıt" (Deploy) > "Dağıtımları Yönet" (Manage Deployments)
 * 3. Kalem (Düzenle) ikonuna basıp Sürüm kısmını "Yeni Sürüm" (New Version) seçip "Dağıt" (Deploy) deyiniz.
 * 
 * KRİTİK DÜZELTME:
 * - Düzenleme (update) işleminde var olan satır DOĞRUDAN o satır üzerinde güncellenir.
 * - Düzenleme işleminde ASLA yeni bir satır (appendRow) açılmaz!
 * - Sıra No eşleşmesi hem sayı (numeric) hem metin (string) olarak yapılır.
 * - Silme (delete) işleminde ilgili satır tablodan tamamen silinir.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000); // 15 saniye kilit bekleme süresi

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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

    // Satırları tara (Sıra No = Sütun 1 / A)
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // Yardımcı fonksiyon: Sıra No'ya göre satır indeksini bul (1-tabanlı)
    function findRowIndex() {
      if (!targetId && isNaN(targetNum)) return -1;
      for (var r = 1; r < values.length; r++) { // Başlık satırını atla (r=0)
        var cellVal = values[r][0];
        if (cellVal === undefined || cellVal === null || cellVal === "") continue;
        
        var cellStr = cellVal.toString().trim();
        var cellNum = parseInt(cellStr, 10);

        // Hem string hem sayı olarak karşılaştır
        if (cellStr === targetId || (!isNaN(targetNum) && !isNaN(cellNum) && cellNum === targetNum)) {
          return r + 1; // getRange 1-tabanlıdır
        }
      }
      return -1;
    }

    // ==========================================
    // 1. İŞLEM: GÜNCELLEME (UPDATE)
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
        // ASLA YENİ SATIR AÇILMAZ! Hata dönülerek mevcut kayıtlar korunur
        return createJsonResponse({ 
          status: "error", 
          action: "update", 
          id: targetId, 
          message: "HATA: Sıra No " + targetId + " e-tabloda bulunamadı! Yeni satır açılmadı." 
        });
      }

      var item = body.data || body;
      var updatedRowValues = [
        targetId,
        item.tarih || "",
        item.kaptanPilot || "",
        item.ikinciPilot || "",
        item.teknisyen1 || "",
        item.operator1 || "",
        item.teknisyen2 || "",
        item.operator2 || "",
        item.gorevTipi || "",
        item.gorevBolgesi || "",
        item.kalkis || "",
        item.inis || "",
        item.ucusSuresi || "",
        item.k9YanginHektar !== undefined ? item.k9YanginHektar : "",
        item.miktarCekim !== undefined ? item.miktarCekim : "",
        item.tk9GorevHektar !== undefined ? item.tk9GorevHektar : "",
        item.uyduDk !== undefined ? item.uyduDk : "",
        item.aciklama || ""
      ];

      // İlgili satırı doğrudan yerinde güncelle (1. sütundan itibaren 18 sütun)
      sheet.getRange(rowToUpdate, 1, 1, updatedRowValues.length).setValues([updatedRowValues]);

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
      
      // En yüksek mevcut Sıra No'yu hesapla
      var maxId = 0;
      for (var m = 1; m < values.length; m++) {
        var valNum = parseInt(values[m][0], 10);
        if (!isNaN(valNum) && valNum > maxId) {
          maxId = valNum;
        }
      }
      var nextId = (addData.id || (maxId > 0 ? maxId + 1 : values.length)).toString();

      var newRow = [
        nextId,
        addData.tarih || "",
        addData.kaptanPilot || "",
        addData.ikinciPilot || "",
        addData.teknisyen1 || "",
        addData.operator1 || "",
        addData.teknisyen2 || "",
        addData.operator2 || "",
        addData.gorevTipi || "",
        addData.gorevBolgesi || "",
        addData.kalkis || "",
        addData.inis || "",
        addData.ucusSuresi || "",
        addData.k9YanginHektar !== undefined ? addData.k9YanginHektar : "",
        addData.miktarCekim !== undefined ? addData.miktarCekim : "",
        addData.tk9GorevHektar !== undefined ? addData.tk9GorevHektar : "",
        addData.uyduDk !== undefined ? addData.uyduDk : "",
        addData.aciklama || ""
      ];

      sheet.appendRow(newRow);
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
