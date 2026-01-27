import {
    loadTemplate,
    insertDataIntoSheet,
    fetchDataFromSheet,
    saveSpreadsheet
} from './helpers/ironcalc.js';

/**
 * Process spreadsheet calculation
 * Mirrors the Google Sheets POC flow:
 * 1. Load template (instead of copying Google Sheet)
 * 2. Insert input data
 * 3. Evaluate formulas (automatic in IronCalc after setUserInput + evaluate())
 * 4. Fetch calculated results
 * 5. Save to XLSX file
 */
const processCalculation = async (counter, inputData = {}) => {
    console.log("******************START*******************");
    console.time(`totalProcessTime_${counter}`);

    try {
        // Step 1: Load template
        console.time(`loadTemplate_${counter}`);
        const model = loadTemplate(`CALCULATION_${counter}`);
        console.timeEnd(`loadTemplate_${counter}`);

        // Step 2: Insert data and evaluate
        console.time(`insertAndEvaluate_${counter}`);
        insertDataIntoSheet(model, inputData);
        console.timeEnd(`insertAndEvaluate_${counter}`);

        // Step 3: Fetch calculated results
        console.time(`fetchResults_${counter}`);
        const results = fetchDataFromSheet(model);
        console.timeEnd(`fetchResults_${counter}`);

        // Step 4: Save to file
        console.time(`saveFile_${counter}`);
        const outputPath = saveSpreadsheet(model);
        console.timeEnd(`saveFile_${counter}`);

        console.timeEnd(`totalProcessTime_${counter}`);
        console.log("******************END*******************\n");

        return {
            success: true,
            results,
            outputPath
        };

    } catch (error) {
        console.error(`Error in calculation ${counter}:`, error.message);
        console.timeEnd(`totalProcessTime_${counter}`);
        console.log("******************END (ERROR)*******************\n");

        return {
            success: false,
            error: error.message
        };
    }
};

// Main execution
const main = async () => {
    console.log("=== IronCalc POC - Local Spreadsheet Engine ===\n");

    // Test with default data (same as Google Sheets POC)
    const result = await processCalculation(1);

    if (result.success) {
        console.log("\n=== CALCULATION RESULTS ===");
        console.log(JSON.stringify(result.results, null, 2));
        console.log(`\nOutput saved to: ${result.outputPath}`);
    }

    // Example: Test with custom data
    // const customResult = await processCalculation(2, {
    //     rentalValues: ["$4,000", "$3,500", "$3,200", "$4,500", "$3,800", "$3,600", "$4,100", "$5,000"],
    //     grossIncome: "$25,000",
    //     expenseValues: ["$18,000", "$20,000", "$9,000", "$3,000", "$2,000", "$1,000", "$1,500", "$2,500", "$1,000"]
    // });

    console.log("\n=== PROCESS COMPLETED! ===");
};

main().catch(console.error);
