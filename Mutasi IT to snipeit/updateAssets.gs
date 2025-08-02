var api_baseurl = 'https://mysnipeit.com/api/v1'; 
var api_bearer_token = 'Bearer mytoken123456';  

// update dari sheet
function updateAssetsFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Result Push Asset");
  var lastRow = sheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var updateFlag = sheet.getRange(i, 8).getValue(); // Kolom H
    var statusColumn = sheet.getRange(i, 9).getValue(); // Kolom I

    // Jika bukan "Yes" atau sudah "Asset updated successfully", lewati baris ini
    if (updateFlag.toLowerCase() !== "yes" || statusColumn === "Asset updated successfully") {
      continue;
    }

    var asset_id = sheet.getRange(i, 1).getValue();
    var asset_tag = sheet.getRange(i, 2).getValue();
    var status_name = sheet.getRange(i, 3).getValue();
    var username = sheet.getRange(i, 4).getValue();
    var location_name = sheet.getRange(i, 5).getValue();
    var notes = sheet.getRange(i, 6).getValue();

    // Asset ID dan Status Name harus ada
    if (asset_id && status_name) {
      var statusID = getStatusIDByName(status_name);
      var userID = username ? getUserIDByUsername(username) : null;
      var locationID = location_name ? getLocationIDByName(location_name) : null;

      if (!statusID) {
        Logger.log("Status ID for '" + status_name + "' not found.");
        sheet.getRange(i, 9).setValue("Update failed: Status ID not found").setBackground('#f4cccc'); // Kolom I
        continue;
      }

      Logger.log("Location ID for '" + location_name + "': " + locationID);
      Logger.log("Status ID for '" + status_name + "': " + statusID);

      // Check-in atau Check-out sesuai status & username
      if (status_name.toLowerCase() === "inbound") {
        checkInAsset(asset_id, username); // Tambahkan username
      } else {
        // Pastikan untuk memeriksa status aset sebelum checkout
        var checkOutResponse = checkOutAsset(asset_id, locationID, userID);
        if (checkOutResponse.status === "failed") {
          Logger.log("Gagal checkout: " + checkOutResponse.message);
          sheet.getRange(i, 9).setValue("Checkout failed: " + checkOutResponse.message).setBackground('#f4cccc');
          continue; // Lanjutkan ke baris berikutnya
        }
      }

      var updateResponse = updateAssetStatus(asset_id, statusID, notes, locationID);

      if (updateResponse && updateResponse.status === "success") {
        sheet.getRange(i, 9).setValue("Asset updated successfully").setBackground('#b6d7a8'); // Kolom I
      } else {
        const errorMessage = updateResponse ? updateResponse.message : "Unknown error";
        sheet.getRange(i, 9).setValue("Update failed: " + errorMessage).setBackground('#f4cccc'); // Kolom I
      }
    } else {
      sheet.getRange(i, 9).setValue("Update failed: Missing asset ID or status").setBackground('#f4cccc'); // Kolom I
    }
  }
}


// fungsi untuk username
function getUserIDByUsername(username) {
  var url = api_baseurl + '/users?search=' + encodeURIComponent(username);
  Logger.log("Fetching user ID for username: " + username);

  var headers = {
    "Authorization": api_bearer_token
  };
  
  var options = {
    "method": "GET",
    "contentType": "application/json",
    "headers": headers
  };
  
  try {
    var response = JSON.parse(UrlFetchApp.fetch(url, options));
    Logger.log("Response from API: " + JSON.stringify(response));
    
    var rows = response.rows;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].username === username) {
        Logger.log("Found user ID: " + rows[i].id);
        return rows[i].id;
      }
    }
    Logger.log("User ID not found for username: " + username);
    return null;

  } catch (e) {
    Logger.log("Error fetching user ID: " + e.message);
    return null;
  }
}


// fungsi untuk status
function getStatusIDByName(status_name) {
  var url = api_baseurl + '/statuslabels?search=' + encodeURIComponent(status_name);
  var headers = {
    "Authorization": api_bearer_token
  };
  
  var options = {
    "method": "GET",
    "contentType": "application/json",
    "headers": headers
  };
  
  var response = JSON.parse(UrlFetchApp.fetch(url, options));
  var rows = response.rows;
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row.name === status_name) {
      return row.id;
    }
  }
  return null;
}

// Fungsi untuk ID lokasi berdasarkan nama
function getLocationIDByName(location_name) {
  var url = api_baseurl + '/locations?search=' + encodeURIComponent(location_name);
  var headers = {
    "Authorization": api_bearer_token
  };
  
  var options = {
    "method": "GET",
    "contentType": "application/json",
    "headers": headers
  };
  
  try {
    var response = JSON.parse(UrlFetchApp.fetch(url, options));
    var rows = response.rows;
    
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row.name === location_name) {
        return row.id;
      }
    }
    Logger.log("Location '" + location_name + "' not found.");
    return null; // Tidak ditemukan
  } catch (e) {
    Logger.log("Error fetching location ID: " + e.message);
    return null;
  }
}

// fungsi update lokasi only
function updateAssetLocationOnly(id, locationID, notes) {
  var data = { "notes": notes || "" };

  // tambahkan lokasi
  if (locationID) {
    data["location_id"] = locationID;
  }

  var url = api_baseurl + '/hardware/' + id;
  var headers = {
    "Authorization": api_bearer_token
  };
  
  var options = {
    "method": "PATCH",
    "contentType": "application/json",
    "headers": headers,
    "payload": JSON.stringify(data)
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var jsonResponse = JSON.parse(response.getContentText());
    Logger.log("Update Location Response: " + JSON.stringify(jsonResponse));
    return jsonResponse;

  } catch (e) {
    Logger.log("Error updating location: " + e.message);
    return { status: "failed", message: e.message };
  }
}

// fungsi update status
function updateAssetStatus(id, status, notes, locationID) {
  var data = {
    "status_id": status,
    "notes": notes
  };
  
  if (locationID !== null) {
    data["location_id"] = locationID;
  }

  var url = api_baseurl + '/hardware/' + id;
  var headers = {
    "Authorization": api_bearer_token
  };
  
  var options = {
    "method": "PATCH",
    "contentType": "application/json",
    "headers": headers,
    "payload": JSON.stringify(data)
  };

  try {
    var response = JSON.parse(UrlFetchApp.fetch(url, options));
    Logger.log("Asset status updated: " + JSON.stringify(response));
    return response;

  } catch (e) {
    Logger.log("Error updating asset status: " + e.message);
    return { status: "failed", message: e.message };
  }
}

// fungsi check-in asset
function checkInAsset(asset_id) {
  var url = api_baseurl + '/hardware/' + asset_id + '/checkin';
  var headers = {
    "Authorization": api_bearer_token
  };

  var options = {
    "method": "POST",
    "contentType": "application/json",
    "headers": headers,
    "payload": JSON.stringify({})
  };

  try {
    var response = JSON.parse(UrlFetchApp.fetch(url, options));
    Logger.log("Checkin response: " + JSON.stringify(response));
    return response;
  } catch (e) {
    Logger.log("Error saat checkin: " + e.message);
    return { status: "failed", message: e.message };
  }
}


// fungsi check-out asset
function checkOutAsset(asset_id, locationID, userID) {
  var url = api_baseurl + '/hardware/' + asset_id;
  var headers = {
    "Authorization": api_bearer_token
  };

  // Dapatkan informasi aset untuk memeriksa status checkout
  try {
    var assetDetails = JSON.parse(UrlFetchApp.fetch(url, { method: "GET", headers: headers }));
    var assignedUser = assetDetails.assigned_user; // Pengguna yang saat ini memiliki aset

    // Jika aset sudah tercheckout, lakukan checkin terlebih dahulu
    if (assignedUser) {
      Logger.log("Aset saat ini tercheckout, melakukan checkin terlebih dahulu...");
      var checkinResponse = checkInAsset(asset_id);
      if (checkinResponse.status !== "success") {
        Logger.log("Gagal checkin aset: " + JSON.stringify(checkinResponse));
        return { status: "failed", message: "Gagal checkin sebelum checkout" };
      }
      Logger.log("Checkin berhasil.");
    }
  } catch (e) {
    Logger.log("Error mendapatkan detail aset: " + e.message);
    return { status: "failed", message: e.message };
  }

  // Lanjutkan dengan checkout ke user baru
  var data = {};
  if (locationID) {
    data["assigned_location"] = locationID;
    data["checkout_to_type"] = "location";
  }
  if (userID) {
    data["assigned_user"] = userID;
    data["checkout_to_type"] = "user";
  }

  var checkoutUrl = api_baseurl + '/hardware/' + asset_id + '/checkout';
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "headers": headers,
    "payload": JSON.stringify(data)
  };

  try {
    var checkoutResponse = JSON.parse(UrlFetchApp.fetch(checkoutUrl, options));
    Logger.log("Checkout response: " + JSON.stringify(checkoutResponse));
    return checkoutResponse;
  } catch (e) {
    Logger.log("Error saat checkout: " + e.message);
    return { status: "failed", message: e.message };
  }
}