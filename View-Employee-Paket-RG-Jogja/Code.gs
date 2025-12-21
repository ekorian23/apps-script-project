const SHEET_ID = '1e7L_l3X1v18krmbQtW6POiiFz042FRrfUKJvF3mC3BE';

function doGet() {
  return HtmlService.createTemplateFromFile('index_employee')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('HQ Jogja - Package System View Employee')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Function to get package data for employee view - WITH DEBUG
function getEmployeePackageData() {
  try {
    console.log('=== START getEmployeePackageData ===');
    
    // Buka spreadsheet
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    console.log('Spreadsheet opened:', spreadsheet.getName());
    
    // Cek semua sheet yang ada
    const sheets = spreadsheet.getSheets();
    console.log('Available sheets:', sheets.map(s => s.getName()));
    
    // Ambil sheet View Only
    const sheet = spreadsheet.getSheetByName('View Only');
    if (!sheet) {
      console.error('❌ Sheet "View Only" not found!');
      return [];
    }
    console.log('✅ Sheet "View Only" found');
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    console.log('Sheet dimensions - Last row:', lastRow, 'Last column:', lastCol);
    
    if (lastRow < 2) {
      console.log('ℹ️ No data found (lastRow < 2)');
      return [];
    }
    
    // Ambil data dari kolom A, B, C, D
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    console.log('Raw data retrieved:', data.length, 'rows');
    
    // Debug: Tampilkan 5 data pertama
    console.log('First 5 rows of data:');
    for (let i = 0; i < Math.min(5, data.length); i++) {
      console.log(`Row ${i + 2}:`, {
        nama: data[i][0],
        tanggal: data[i][1],
        ekspedisi: data[i][2],
        status: data[i][3],
        types: {
          nama: typeof data[i][0],
          tanggal: typeof data[i][1],
          ekspedisi: typeof data[i][2],
          status: typeof data[i][3]
        }
      });
    }
    
    // REVERSE DATA: Data terbaru (paling bawah di sheet) jadi paling atas di web
    const reversedData = data.reverse();
    
    const packageData = reversedData
      .filter(row => row[0] && row[0].toString().trim() !== '') // Filter baris dengan nama kosong
      .map((row, index) => {
        const namaPenerima = row[0] ? row[0].toString().trim() : '';
        const tanggalDiterima = row[1];
        const ekspedisi = row[2] ? row[2].toString().trim() : '';
        const status = row[3] ? row[3].toString().trim() : '';
        
        // Handle tanggal - coba berbagai format
        let formattedTanggalDiterima = '';
        try {
          if (tanggalDiterima instanceof Date) {
            formattedTanggalDiterima = Utilities.formatDate(tanggalDiterima, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          } else if (tanggalDiterima) {
            // Coba parse sebagai string date
            const dateObj = new Date(tanggalDiterima);
            if (!isNaN(dateObj.getTime())) {
              formattedTanggalDiterima = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
            } else {
              formattedTanggalDiterima = tanggalDiterima.toString();
            }
          }
        } catch (e) {
          formattedTanggalDiterima = 'Invalid Date';
          console.error('Date formatting error:', e);
        }
        
        // Determine status text
        let statusText = 'Pending';
        if (status && status.toString().toLowerCase() === 'done') {
          statusText = 'Done';
        }
        
        return {
          namaPenerima: namaPenerima,
          tanggalDiterima: formattedTanggalDiterima,
          ekspedisi: ekspedisi,
          status: statusText,
          timestamp: tanggalDiterima instanceof Date ? tanggalDiterima.getTime() : 0,
          isPending: statusText === 'Pending'
        };
      });
    
    console.log('✅ Processed employee package data:', packageData.length);
    console.log('Sample processed data:', packageData.slice(0, 3));
    console.log('=== END getEmployeePackageData ===');
    
    return packageData;
    
  } catch (error) {
    console.error('❌ Error getting employee package data:', error);
    console.error('Error stack:', error.stack);
    return [];
  }
}

// Function untuk test manual
function testDataRetrieval() {
  try {
    const result = getEmployeePackageData();
    return {
      success: true,
      dataCount: result.length,
      sampleData: result.slice(0, 3),
      message: `Retrieved ${result.length} packages`
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      message: 'Failed to retrieve data'
    };
  }
}