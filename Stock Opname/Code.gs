function doGet(request) {
  var template = HtmlService.createTemplateFromFile('index') 
  var pageData = template.evaluate()
  .setTitle("Form Stock Opname")
  .addMetaTag("viewport","width=device-width, initial-scale=1")
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  return pageData;
}

//mengambil data employee dari sheet
function getEmployeeNames() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data_employee');
  var range = sheet.getRange('A2:A');
  var names = range.getValues().flat().filter(String);
  return names;
}

//mengambil data employee dari sheet
function getEmployeeDetails(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data_employee');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      return {
        namaLengkap: data[i][0],
        nip: data[i][1],
        posisi: data[i][2],
        departemen: data[i][3],
        lokasi: data[i][4],
        email: data[i][5]
      };
    }
  }
  return {};
}

//mengambil data asset dari sheet
function getAssetCodes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('raw_data_asset');
  var range = sheet.getRange('A2:A');
  var codes = range.getValues().flat().filter(String);
  return codes;
}

function submitForm(formData) {
  Logger.log("Form Data Received: " + JSON.stringify(formData));

  var employeeDetails = getEmployeeDetails(formData.namaLengkap);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form_Response');
  
  if (!sheet) {
    sheet = ss.insertSheet('Form_Response');
    sheet.appendRow(['Timestamp', 'Nama Lengkap', 'NIP', 'Posisi', 'Departemen', 'Work Location', 'Email', 'Kode Asset', 'Kode Asset Manual', 'Mac Address', 'Serial Number', 'Kode Asset Lainnya']);
  }

  var timestamp = new Date();
  sheet.appendRow([
    timestamp,
    formData.namaLengkap || '',
    employeeDetails.nip || '',
    employeeDetails.posisi || '',
    employeeDetails.departemen || '',
    employeeDetails.lokasi || '',
    employeeDetails.email || '',
    formData.kodeAsset || '',
    formData.kodeAssetManual || '',
    formData.macAddress || '',
    formData.serialNumber || '',
    formData.kodeAssetLainnya || ''
  ]);

  Logger.log("Form data saved successfully");

  return "success";
}
