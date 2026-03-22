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

// Fungsi Login
function doLogin(username, password) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Login');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Login sheet not found'
      };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {
        success: false,
        message: 'No users registered'
      };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const storedUsername = data[i][0];
      const storedPassword = data[i][1];
      
      if (storedUsername && storedPassword) {
        if (storedUsername.toString().trim() === username.trim() && 
            storedPassword.toString().trim() === password.trim()) {
          return {
            success: true,
            message: 'Login successful',
            username: username.trim()
          };
        }
      }
    }
    
    return {
      success: false,
      message: 'Invalid username or password'
    };
    
  } catch (error) {
    console.error('Login Error:', error);
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

// Fungsi Logout
function doLogout() {
  return {
    success: true,
    message: 'Logged out successfully'
  };
}

// Fungsi untuk cek session
function checkSession(sessionUsername) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Login');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Login sheet not found'
      };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {
        success: false,
        message: 'No users registered'
      };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const storedUsername = data[i][0];
      
      if (storedUsername && storedUsername.toString().trim() === sessionUsername.trim()) {
        return {
          success: true,
          message: 'Session valid',
          username: sessionUsername.trim()
        };
      }
    }
    
    return {
      success: false,
      message: 'Session expired or invalid'
    };
    
  } catch (error) {
    console.error('Session Check Error:', error);
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}

function generateID(namaPenerima, ekspedisi, tanggalDiterima) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
  const lastRow = sheet.getLastRow();
  
  const namaDepan = namaPenerima ? namaPenerima.split(' ')[0].toUpperCase() : 'UNKNOWN';
  const ekspedisiCode = ekspedisi ? ekspedisi.substring(0, 3).toUpperCase() : 'EXP';
  
  // Format tanggal menjadi 6 digit: DDMMYY
  let tanggalCode;
  try {
    const date = tanggalDiterima ? new Date(tanggalDiterima) : new Date();
    tanggalCode = Utilities.formatDate(date, Session.getScriptTimeZone(), 'ddMMyy');
  } catch (e) {
    tanggalCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyy');
  }
  
  // Ambil semua ID yang sudah ada
  const existingIDs = new Set();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    data.forEach(row => {
      if (row[0] && row[0].toString().trim() !== '') {
        existingIDs.add(row[0].toString().trim());
      }
    });
  }
  
  // Cari ID yang unik dengan counter
  let counter = 1;
  let maxAttempts = 999;
  
  while (counter <= maxAttempts) {
    const newID = `RG-${namaDepan}-${ekspedisiCode}-${tanggalCode}-${counter.toString().padStart(3, '0')}`;
    
    if (!existingIDs.has(newID)) {
      return newID;
    }
    
    counter++;
  }
  
  const timestamp = new Date().getTime().toString().slice(-6);
  return `RG-${namaDepan}-${ekspedisiCode}-${timestamp}`;
}

function getAvailableIDs() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // 11 kolom (A sampai K)
    
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

// UPDATED: Menambahkan parameter packageType
function saveData(formData, username) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    const finalID = generateID(formData.namaPenerima, formData.ekspedisi, formData.tanggalDiterima);
    
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
    
    const now = new Date();
    const receiveDateTime = new Date(formData.tanggalDiterima);
    
    // UPDATE: Simpan dengan 12 kolom (A sampai L)
    const newRow = [
      finalID,
      formData.namaPenerima,
      now, // Kolom C - timestamp
      formData.ekspedisi, // Kolom D - ekspedisi
      fileUrl, // Kolom E - foto penerimaan
      '',
      '',
      '', // Kolom H - status
      username || '', // Kolom I - username penerima
      '', // Kolom J - username penyerahan
      '', // Kolom K - SLA
      formData.packageType || '' // Kolom L - Package Type (Document/Barang)
    ];
    
    // Cari baris kosong atau tambah di akhir
    let targetRow = lastRow + 1;
    
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

function savePenyerahanData(formData, username) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {
        success: false,
        message: 'No data available to update'
      };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); // 11 kolom (A sampai K)
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
    
    let tanggalDiserahkan;
    if (formData.tanggalDiserahkan) {
      tanggalDiserahkan = new Date(formData.tanggalDiserahkan);
    } else {
      tanggalDiserahkan = new Date();
    }
    
    // Update data
    sheet.getRange(rowIndex, 6).setValue(tanggalDiserahkan);
    sheet.getRange(rowIndex, 7).setValue(fileUrl);
    sheet.getRange(rowIndex, 8).setValue('done');
    sheet.getRange(rowIndex, 10).setValue(username || '');
    
    return {
      success: true,
      message: 'Delivery data saved successfully!',
      id: formData.id,
      tanggalDiserahkan: Utilities.formatDate(tanggalDiserahkan, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
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
    
    if (!dashboardSheet) {
      return getEmptyDashboardData();
    }
    
    const lastRow = dashboardSheet.getLastRow();
    
    if (lastRow < 2) {
      return getEmptyDashboardData();
    }
    
    const dashboardValues = dashboardSheet.getRange('A2:F2').getValues()[0];
    
    const totalPaket = dashboardValues[0] || 0;
    const paketSelesai = dashboardValues[1] || 0;
    const paketBelumSelesai = dashboardValues[2] || 0;
    const overSLA = dashboardValues[3] || 0;
    const pendingIDsText = dashboardValues[5] || '';
    
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

// FIXED: Ambil data dari kolom A sampai L (12 kolom) termasuk Package Type
function getAllPackageData() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('MainData');
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return [];
    }
    
    // Ambil data dari kolom A sampai L (12 kolom)
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    // REVERSE DATA: Data terbaru (paling bawah di sheet) jadi paling atas di table
    const reversedData = data.reverse();
    
    const packageData = reversedData
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map((row, index) => {
        const id = row[0] || '';
        const packageType = row[11] || ''; // Kolom L - Package Type
        const namaPenerima = row[1] || '';
        const tanggalDiterima = row[2] ? new Date(row[2]) : null;
        const ekspedisi = row[3] || '';
        const fotoPenerimaan = convertToDirectImageUrl(row[4] || '');
        const tanggalDiserahkan = row[5] ? new Date(row[5]) : null;
        const fotoPenyerahan = convertToDirectImageUrl(row[6] || '');
        const status = row[7] || '';
        const usernamePenerima = row[8] || '';
        const usernamePenyerahan = row[9] || '';
        const slaStatusRaw = row[10] || '';
        
        const formattedTanggalDiterima = tanggalDiterima ? 
          Utilities.formatDate(tanggalDiterima, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '';
        
        const formattedTanggalDiserahkan = tanggalDiserahkan ? 
          Utilities.formatDate(tanggalDiserahkan, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '';
        
        let statusText = 'Pending';
        if (status === 'done') {
          statusText = 'Done';
        }
        
        let slaStatusText = slaStatusRaw.toString().trim();

        // Override SLA status: jika masih pending dan sudah > 3 hari, set Over SLA
        if (status !== 'done' && tanggalDiterima) {
          const now = new Date();
          const diffMs = now - tanggalDiterima;
          const diffDays = diffMs / (1000 * 60 * 60 * 24);
          if (diffDays > 3) {
            slaStatusText = 'Over SLA';
          }
        }
                
        const originalIndex = data.length - 1 - index;
        const actualRow = originalIndex + 2;
        
        return {
          id: id,
          packageType: packageType, // Package Type dari kolom L
          namaPenerima: namaPenerima,
          tanggalDiterima: formattedTanggalDiterima,
          ekspedisi: ekspedisi,
          fotoPenerimaan: fotoPenerimaan,
          tanggalDiserahkan: formattedTanggalDiserahkan,
          fotoPenyerahan: fotoPenyerahan,
          status: statusText,
          slaStatus: slaStatusText,
          usernamePenerima: usernamePenerima,
          usernamePenyerahan: usernamePenyerahan,
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

function convertToDirectImageUrl(url) {
  if (!url || url.trim() === '') return '';
  
  try {
    if (url.includes('drive.google.com/uc?id=') || url.includes('lh3.googleusercontent.com')) {
      return url;
    }
    
    let fileId = '';
    
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
      return `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;
    }
    
    return url;
  } catch (error) {
    console.error('Error converting Google Drive URL:', error);
    return url;
  }
}
