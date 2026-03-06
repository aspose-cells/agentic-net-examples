using Aspose.Cells;
using System;

class ApplySharedFormula
{
    static void Main()
    {
        // Load the workbook without parsing formulas on open (faster for large files)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet and its cells collection
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Apply a shared formula to column B (B1:B100) that doubles the value in column A
        // Formula: =A1*2, rows = 100, columns = 1
        cells["B1"].SetSharedFormula("=A1*2", 100, 1);

        // Recalculate all formulas so that the new shared formula values are materialized
        workbook.CalculateFormula();

        // Save the workbook with the applied shared formula
        workbook.Save("output.xlsx");
    }
}