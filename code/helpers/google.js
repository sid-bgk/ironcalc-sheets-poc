const { google } = require('googleapis');
const fs = require('fs');

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
  ]
});

const drive = google.drive({ version: "v3", auth });
const sheets = google.sheets({ version: "v4", auth });

const templateSheetID = "1zXhXvwCnAWm03VE1NqGIK95b907_E8fAS7kW2779uLo";

const uploadInitialSheet = async (sheetName) => {
    const newSheetName = `${sheetName}_${Date.now()}`;

    // const fileMetadata = {
    //     name: newSheetName,
    //     mimeType: "application/vnd.google-apps.spreadsheet",
    //     parents: ["********************"]
    // };

    // const media = {
    //     mimeType:
    //     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    //     body: fs.createReadStream(__dirname + "/./../../documents/template.xlsx")
    // };

    // const uploadRes = await drive.files.create({
    //     requestBody: fileMetadata,
    //     media,
    //     fields: "id",
    //     supportsAllDrives: true
    // });

    // const spreadsheetId = uploadRes.data.id;
    // console.log("Created Sheet:", spreadsheetId);

    // return spreadsheetId;

    const copyRes = await drive.files.copy({
        fileId: templateSheetID,
        supportsAllDrives: true,
        requestBody: {
            name: newSheetName,
            parents: ["0AOT_Ntwmd1jUUk9PVA"],
            // mimeType: 'application/vnd.google-apps.spreadsheet'
        }
    });

    return copyRes.data.id;
}

const insertDataIntoSheet = async (spreadsheetId) => {
    await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        valueInputOption: "USER_ENTERED",
        requestBody: {
            data: [
                {
                    range: "RentalCalc!H6:H13",
                    values: [["$3,200"], ["$2,700"],["$2,850"],["$3,500"],["$2,800"],["$2,900"],["$3,100"],["$4,200"]]
                },
                {
                    range: "DscrCalc!B3",
                    values: [["$21,155"]]
                },
                {
                    range: "DscrCalc!B5:B13",
                    values: [["$16,000"], ["$18,500"], ["$8,500"], ["$2,500"], ["$1,800"], ["$800"], ["$1,200"], ["$2,100"], ["$750"]]
                }
            ]
        }
    });
}

const fetchDataFromSheet = async (spreadsheetId) => {
    const res = await sheets.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges: [
            "DscrCalc!G12",
            "DscrCalc!G17",
            "DscrCalc!G29",
            "DscrCalc!G31",
            "DscrCalc!G34",
        ],
        valueRenderOption: "UNFORMATTED_VALUE"
    });

    // console.log("RESPONSE: ", res.data.valueRanges)

    res.data.valueRanges.forEach(r => {
        console.log(r.range, r.values);
    });
}

const deleteSpreadsheetByID = async spreadsheetId => {
    await drive.files.delete({
      fileId: spreadsheetId,
      supportsAllDrives: true,
    });
}

const downloadSpreadsheet = async spreadsheetId => {
    const exportRes = await drive.files.export(
        {
        fileId: spreadsheetId,
        mimeType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        },
        { responseType: "stream" }
    );

    const out = fs.createWriteStream(__dirname + `/./../../result/${spreadsheetId}.xlsx`);

    await new Promise((resolve, reject) => {
        exportRes.data
        .pipe(out)
        .on("finish", resolve)
        .on("error", reject);
    });
}

module.exports = {
    uploadInitialSheet,
    insertDataIntoSheet,
    deleteSpreadsheetByID,
    downloadSpreadsheet,
    fetchDataFromSheet
};