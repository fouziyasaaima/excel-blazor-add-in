export async function GetData()
{ 
    var excelData;
    var ActiveCellFormula;
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getUsedRange();
        let activeCell = context.workbook.getActiveCell();
        activeCell.load("formulas");
        range.load("values");
        await context.sync();
        excelData = JSON.stringify(range.values, null, 4);
        ActiveCellFormula = JSON.stringify(activeCell.formulas, null, 4);
        console.log(JSON.stringify(range.values, null, 4));
    });

    return "ActiveCell Formula : " + ActiveCellFormula + "Sheet Data: " + excelData ;
   
    
}


