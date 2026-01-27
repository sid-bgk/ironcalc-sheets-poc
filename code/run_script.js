const { insertDataIntoSheet, uploadInitialSheet, downloadSpreadsheet, deleteSpreadsheetByID, fetchDataFromSheet } =  require("./helpers/google");


const createSheetAndAddData = async (counter) => {
    console.log("******************START*******************");
    console.time(`sheetProcessStarted_${counter}`);

    console.time(`uploadTemplateSheet_${counter}`);
    const sheetID = await uploadInitialSheet(`TESTING_CONCURRENCY_${counter + 1}`)
    console.timeEnd(`uploadTemplateSheet_${counter}`);
    
    console.time(`insertDataIntoTemplateSheet_${counter}`);
    await insertDataIntoSheet(sheetID);
    console.timeEnd(`insertDataIntoTemplateSheet_${counter}`);

    await fetchDataFromSheet(sheetID)
    
    // await new Promise((r) => setTimeout(r, 2000));
    await downloadSpreadsheet(sheetID);
    
    // await new Promise(r => setTimeout(r, 1500));
    // await deleteSpreadsheetByID(sheetID);
    console.timeEnd(`sheetProcessStarted_${counter}`);
    console.log("******************END*******************\n\n\n\n");
}

for (let i = 0; i < 1; i++) {
    createSheetAndAddData(i + 1);
}

setTimeout(() => {
    console.log("PROCESS COMPLETED!")
}, 15000)