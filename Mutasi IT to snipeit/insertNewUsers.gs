function insertNewUsers() {

    var api_baseurl = 'https://mysnipeit.com/api/v1';
    var api_bearer_token = 'Bearer mytoken123456';

    var sheet = SpreadsheetApp.getActive().getSheetByName('Result Push User');
    var data = sheet.getDataRange().getValues();

    for (var i = 2; i <= data.length; i++) {
        var row = data[i - 1];
        var updateFlag = (row[8] || "").trim().toLowerCase(); // Kolom I
        var statusColumn = (row[9] || "").trim(); // Kolom J

        // Skip jika kolom I bukan "yes" atau sudah sukses
        if (updateFlag !== "yes" || statusColumn === "User was successfully created") {
            continue;
        }

        let first_name = row[0];  // Kolom A
        let last_name = row[1];   // Kolom B
        let username = row[2];    // Kolom C
        let employee_number = String(row[3]).trim(); // Kolom D
        let email = row[4];       // Kolom E
        let password = row[6];    // Kolom G
        let notes = row[7];       // Kolom H
        let location_id = String(row[10] || "").trim();

        // Validasi employee_number
        if (!/^(?:NIP|NIP2)\.\d{2}\.\d{2}\.\d{2}\.\d{5}$/.test(employee_number)) {
            Logger.log('Invalid Employee Number Format: ' + employee_number);
            sheet.getRange(i, 9).setValue("Invalid Employee Number Format").setBackground('#f4cccc');
            continue;
        }

        // Persiapkan payload
        let url = api_baseurl + '/users';
        var payloadData = {
            first_name: first_name,
            last_name: last_name,
            username: username,
            employee_num: employee_number,
            email: email,
            password: password,
            password_confirmation: password,
            notes: notes,
        };

        // Jika ID lokasi ada, tambahkan ke payload
        if (location_id) {
            payloadData.location_id = location_id;
        }

        var options = {
            method: 'post',
            muteHttpExceptions: true,
            headers: {
                Accept: 'application/json',
                Authorization: api_bearer_token,
                'Content-Type': 'application/json',
            },
            payload: JSON.stringify(payloadData),
        };

        // Kirim ke API
        var response = UrlFetchApp.fetch(url, options);
        var response_msg = JSON.parse(response.getContentText());

        Logger.log('Payload sent: ' + options.payload);
        Logger.log('Response Text: ' + response.getContentText());

        if (response.getResponseCode() === 200 && response_msg.status === 'success') {
            sheet.getRange(i, 9).setValue("User was successfully created").setBackground('#b6d7a8');
        } else {
            const errorMessage = parseValidationErrors(response_msg.messages || ["Unknown error"]);
            sheet.getRange(i, 9).setValue("Update failed: " + errorMessage).setBackground('#f4cccc');
        }

        Utilities.sleep(1000); // Untuk menghindari rate-limit API
    }
}

function parseValidationErrors(messages) {
    if (!Array.isArray(messages)) {
        if (typeof messages === 'string') {
            return messages;
        } else if (typeof messages === 'object') {
            return Object.values(messages).flat().join(', ');
        } else {
            return "Unknown error";
        }
    }
    return messages.join(', ');

}
