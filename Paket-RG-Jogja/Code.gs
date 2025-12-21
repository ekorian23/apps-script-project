const SHEET_ID = '1fkolUlyfJHPYExyG6HCCmAta0KF7d8J1mwRKGOXNlFY';
const DRIVE_FOLDER_ID = '1QSl2G8MRtHihPBV6yXmEIAjSOy0g1_jH';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('HQ RG Jogja Package System')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function generateID(namaPenerima, ekspedisi) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
  const lastRow = sheet.getLastRow();
  
  const namaDepan = namaPenerima ? namaPenerima.split(' ')[0].toUpperCase() : 'UNKNOWN';
  const ekspedisiCode = ekspedisi ? ekspedisi.substring(0, 3).toUpperCase() : 'EXP';
  
  let counter = 1;
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    
    const filteredData = data.filter(row => {
      const existingID = row[0];
      if (!existingID) return false;
      
      const idParts = existingID.split('-');
      if (idParts.length !== 4) return false;
      
      const existingNama = idParts[1];
      const existingEkspedisi = idParts[2];
      
      return existingNama === namaDepan && existingEkspedisi === ekspedisiCode;
    });
    
    counter = filteredData.length + 1;
  }
  
  return `RG-${namaDepan}-${ekspedisiCode}-${counter.toString().padStart(3, '0')}`;
}

function getAvailableIDs() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    const availableIDs = data.filter(row => {
      const id = row[0];
      const statusPenyerahan = row[7];
      return id && !statusPenyerahan;
    }).map(row => row[0]);
    
    return availableIDs;
  } catch (error) {
    return [];
  }
}

function saveData(formData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    const finalID = generateID(formData.namaPenerima, formData.ekspedisi);
    
    let fileUrl = '';
    if (formData.documentation) {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(formData.documentation.split(',')[1]),
        formData.fileType,
        formData.fileName
      );
      
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }
    
    const newRow = [
      finalID,
      formData.namaPenerima,
      new Date(formData.tanggalDiterima),
      formData.ekspedisi,
      fileUrl,
      '',
      '',
      ''
    ];
    
    // Cari baris kosong atau tambah di akhir
    let targetRow = lastRow + 1;
    
    // Cek jika ada baris kosong di antara data
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < data.length; i++) {
        if (!data[i][0] || data[i][0] === '') {
          targetRow = i + 2;
          break;
        }
      }
    }
    
    sheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
    
    return {
      success: true,
      message: 'Receive data saved successfully!',
      id: finalID
    };
    
  } catch (error) {
    console.error('Save Data Error:', error);
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

function savePenyerahanData(formData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {
        success: false,
        message: 'No data available to update'
      };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    let rowIndex = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === formData.id) {
        rowIndex = i + 2;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {
        success: false,
        message: 'ID not found: ' + formData.id
      };
    }
    
    // Cek apakah sudah diserahkan
    const statusCell = sheet.getRange(rowIndex, 8).getValue();
    if (statusCell === 'done') {
      return {
        success: false,
        message: 'Package already delivered'
      };
    }
    
    let fileUrl = '';
    if (formData.dokumentasiPenyerahan) {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(formData.dokumentasiPenyerahan.split(',')[1]),
        formData.fileType,
        formData.fileName
      );
      
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }
    
    // Update data
    sheet.getRange(rowIndex, 6).setValue(new Date(formData.tanggalDiserahkan));
    sheet.getRange(rowIndex, 7).setValue(fileUrl);
    sheet.getRange(rowIndex, 8).setValue('done');
    
    return {
      success: true,
      message: 'Delivery data saved successfully!',
      id: formData.id
    };
    
  } catch (error) {
    console.error('Save Penyerahan Error:', error);
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

function getDashboardData() {
  try {
    const dashboardSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Dashboard');
    
    // Cek jika sheet Dashboard ada dan memiliki rumus
    if (!dashboardSheet) {
      return getEmptyDashboardData();
    }
    
    const lastRow = dashboardSheet.getLastRow();
    
    if (lastRow < 2) {
      return getEmptyDashboardData();
    }
    
    // Ambil data dari hasil rumus ARRAYFORMULA di A2:F2
    const dashboardValues = dashboardSheet.getRange('A2:F2').getValues()[0];
    
    const totalPaket = dashboardValues[0] || 0;
    const paketSelesai = dashboardValues[1] || 0;
    const paketBelumSelesai = dashboardValues[2] || 0;
    const overSLA = dashboardValues[3] || 0;
    const pendingIDsText = dashboardValues[5] || '';
    
    // Convert pending IDs text ke array
    let idBelumSelesai = [];
    if (pendingIDsText && pendingIDsText.toString().trim() !== '') {
      idBelumSelesai = pendingIDsText.toString().split(',').map(id => id.trim()).filter(id => id !== '');
    }
    
    return {
      totalPaket: totalPaket,
      paketSelesai: paketSelesai,
      paketBelumSelesai: paketBelumSelesai,
      overSLA: overSLA,
      idBelumSelesai: idBelumSelesai
    };
    
  } catch (error) {
    console.error('Error getting dashboard data:', error);
    return getEmptyDashboardData();
  }
}

function getEmptyDashboardData() {
  return {
    totalPaket: 0,
    paketSelesai: 0,
    paketBelumSelesai: 0,
    overSLA: 0,
    idBelumSelesai: []
  };
}

// NEW FUNCTION: Get all package data for table - UPDATED WITH REVERSE
function getAllPackageData() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return [];
    }
    
    // Ambil data dari kolom A sampai I (9 kolom)
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    
    // REVERSE DATA: Data terbaru (paling bawah di sheet) jadi paling atas di table
    const reversedData = data.reverse();
    
    const packageData = reversedData
      .filter(row => row[0] && row[0].toString().trim() !== '') // Filter baris dengan ID kosong
      .map((row, index) => {
        const id = row[0] || '';
        const namaPenerima = row[1] || '';
        const tanggalDiterima = row[2] ? new Date(row[2]) : null;
        const ekspedisi = row[3] || '';
        const fotoPenerimaan = convertToDirectImageUrl(row[4] || ''); // Convert URL
        const tanggalDiserahkan = row[5] ? new Date(row[5]) : null;
        const fotoPenyerahan = convertToDirectImageUrl(row[6] || ''); // Convert URL
        const status = row[7] || '';
        const slaStatus = row[8] || '';
        
        // Format dates
        const formattedTanggalDiterima = tanggalDiterima ? 
          Utilities.formatDate(tanggalDiterima, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
        
        const formattedTanggalDiserahkan = tanggalDiserahkan ? 
          Utilities.formatDate(tanggalDiserahkan, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
        
        // Determine status text
        let statusText = 'Pending';
        if (status === 'done') {
          statusText = 'Done';
        }
        
        // Hitung row asli di spreadsheet (karena data di-reverse)
        const originalIndex = data.length - 1 - index;
        const actualRow = originalIndex + 2; // +2 karena mulai dari row 2
        
        return {
          id: id,
          namaPenerima: namaPenerima,
          tanggalDiterima: formattedTanggalDiterima,
          ekspedisi: ekspedisi,
          fotoPenerimaan: fotoPenerimaan,
          tanggalDiserahkan: formattedTanggalDiserahkan,
          fotoPenyerahan: fotoPenyerahan,
          status: statusText,
          slaStatus: slaStatus,
          rowIndex: actualRow,
          timestamp: tanggalDiterima ? tanggalDiterima.getTime() : 0
        };
      });
    
    console.log('Processed package data:', packageData.length);
    return packageData;
    
  } catch (error) {
    console.error('Error getting package data:', error);
    return [];
  }
}

// NEW FUNCTION: Convert Google Drive view URL to direct image URL
function convertToDirectImageUrl(url) {
  if (!url || url.trim() === '') return '';
  
  try {
    // Jika sudah direct URL, return as is
    if (url.includes('drive.google.com/uc?id=') || url.includes('lh3.googleusercontent.com')) {
      return url;
    }
    
    // Extract file ID dari Google Drive URL
    let fileId = '';
    
    // Pattern untuk berbagai format Google Drive URL
    const patterns = [
      /\/file\/d\/([a-zA-Z0-9_-]+)/,
      /id=([a-zA-Z0-9_-]+)/,
      /\/d\/([a-zA-Z0-9_-]+)\//,
      /([a-zA-Z0-9_-]{25,})/
    ];
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match && match[1]) {
        fileId = match[1];
        break;
      }
    }
    
    if (fileId) {
      // Return direct image URL untuk thumbnail/preview
      return `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;
    }
    
    return url; // Return original URL jika tidak bisa extract
  } catch (error) {
    console.error('Error converting Google Drive URL:', error);
    return url; // Return original URL jika error
  }
}