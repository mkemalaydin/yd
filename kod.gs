const SHEET_NAME = "Veritabani";

// 1. Veritabanı sekmesini otomatik oluşturan kurulum fonksiyonu
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Sütun Başlıkları
    sheet.appendRow(["Kurum/Birim ID", "Değerlendirme Tarihi", "Program Adı", "Toplam Puan", "Durum", "Genel Özet", "Güçlü Yönler", "Gelişmeye Açık Yönler", "Detay_JSON"]);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#f8fafc");
    sheet.setColumnWidth(9, 500); // JSON verisi uzun olacağı için sütunu genişletiyoruz
  }
}

// 2. OPTIONS metodu: Tarayıcının CORS (Preflight) kontrolünü geçmek için zorunludur.
// Özellikle büyük veri (uzun metin/görsel) gönderildiğinde tarayıcı önce bu fonksiyonu çağırır.
function doOptions(e) {
  return createCORSResponse("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function createCORSResponse(content) {
  return ContentService.createTextOutput(content);
}

function handleFileUploadRequest(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createCORSResponse(JSON.stringify({status: "error", message: "Dosya verisi alınamadı."}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const payload = JSON.parse(e.postData.contents);
    const fileName = String(payload.fileName || "uploaded-file");
    const mimeType = String(payload.mimeType || "application/octet-stream");
    const base64Data = String(payload.data || "");

    if (!base64Data) {
      return createCORSResponse(JSON.stringify({status: "error", message: "Boş dosya verisi gönderildi."}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return createCORSResponse(JSON.stringify({status: "success", url: file.getUrl(), title: file.getName()}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return createCORSResponse(JSON.stringify({status: "error", message: "Dosya yükleme hatası: " + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function compressStringForStorage(text) {
  const value = String(text || '');
  if (!value) return value;
  const gzippedBlob = Utilities.gzip(Utilities.newBlob(value, 'application/json'));
  return '#GZ#' + Utilities.base64Encode(gzippedBlob.getBytes());
}

function decompressStringFromStorage(value) {
  if (typeof value !== 'string' || !value.startsWith('#GZ#')) return value;
  const decodedBytes = Utilities.base64Decode(value.slice(4));
  const blob = Utilities.newBlob(decodedBytes, 'application/octet-stream');
  const decompressed = Utilities.ungzip(blob);
  return decompressed.getDataAsString();
}

// 3. Frontend'den gelen verileri kaydetme (POST) işlemi
function doPost(e) {
  try {
    // Eğer postData boşsa güvenli çıkış yap
    if (!e || !e.postData || !e.postData.contents) {
      return createCORSResponse(JSON.stringify({status: "error", message: "Veri alınamadı veya boş gönderildi."}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const action = e.parameter && e.parameter.action;
    if (action === 'uploadFile') {
      return handleFileUploadRequest(e);
    }

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

    const id = data.projectNumber;
    const timestamp = new Date();
    
    // Detaylar objesini string'e çevir
    const detailsJsonString = JSON.stringify(data.details);
    let storedDetails = detailsJsonString;

    // Google Sheets hücre limiti kontrolü (Maks: 50.000 karakter)
    // Detay JSON'u tek hücrede saklandığı için uzun veri doğrudan sınırı aşabilir.
    if (detailsJsonString.length > 50000) {
        const compressed = compressStringForStorage(detailsJsonString);
        if (compressed.length < 50000) {
            storedDetails = compressed;
        } else {
            return createCORSResponse(JSON.stringify({
                status: "error",
                message: `Veri boyutu çok büyük! Detay JSON uzunluğu ${detailsJsonString.length} karakter. Sıkıştırılmış hali ${compressed.length} karakter ve yine sınırı aşıyor. Lütfen yorumları kısaltın veya sayfayı daha az metinle kaydedin.`
            }))
            .setMimeType(ContentService.MimeType.JSON);
        }
    }

    // Satıra yazılacak veri dizisi
    const rowData = [
      id,
      timestamp,
      data.projectName,
      data.grandTotal,
      data.status,
      data.overall.summary,
      data.overall.strengths,
      data.overall.weaknesses,
      storedDetails
    ];

    // ID daha önce kaydedilmiş mi kontrol et (Güncelleme mantığı)
    const dataRange = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] == id) {
        foundRow = i + 1; // Apps script satırları 1'den başlar
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]); // Var olanı güncelle
    } else {
      sheet.appendRow(rowData); // Yeni kayıt ekle
    }

    return createCORSResponse(JSON.stringify({status: "success", message: "Veritabanına başarıyla kaydedildi!"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return createCORSResponse(JSON.stringify({status: "error", message: "Sunucu Hatası: " + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 4. Frontend'e verileri ve ID listesini gönderme (GET) işlemi
function doGet(e) {
  // Preflight veya parametresiz durumlar için güvenlik
  if (!e || !e.parameter || !e.parameter.action) {
    return createCORSResponse(JSON.stringify({status: "error", message: "Eksik parametre."}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }

  const data = sheet.getDataRange().getValues();
  let response;

  try {
    // Açılır liste (Select) için kayıtlı ID'leri getir
    if (action === "getList") {
      let list = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) list.push(data[i][0]);
      }
      response = JSON.stringify(list);
    }
    // Seçilen ID'nin tüm verilerini (JSON dahil) frontend'e yükle
    else if (action === "load") {
      const projNo = e.parameter.projNo;
      let result = null;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == projNo) {
          let detailsCell = data[i][8] || "{}";
          if (typeof detailsCell === 'string' && detailsCell.startsWith('#GZ#')) {
            detailsCell = decompressStringFromStorage(detailsCell);
          }
          result = {
            projectNumber: data[i][0],
            projectName: data[i][2],
            grandTotal: data[i][3],
            status: data[i][4],
            overall: {
              summary: data[i][5],
              strengths: data[i][6],
              weaknesses: data[i][7]
            },
            details: JSON.parse(detailsCell || "{}")
          };
          break;
        }
      }
      
      if (result) {
        response = JSON.stringify({status: "success", data: result});
      } else {
        response = JSON.stringify({status: "error", message: "Kayıt bulunamadı."});
      }
    } else {
      response = JSON.stringify({status: "error", message: "Geçersiz işlem."});
    }
  } catch(err) {
      response = JSON.stringify({status: "error", message: err.toString()});
  }

  return createCORSResponse(response)
    .setMimeType(ContentService.MimeType.JSON);
}