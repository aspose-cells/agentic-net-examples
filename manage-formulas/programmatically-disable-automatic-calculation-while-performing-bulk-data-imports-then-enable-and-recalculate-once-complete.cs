using System;
using Aspose.Cells;

class BulkImportExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Disable automatic calculation by setting the mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Perform bulk data import (example: 10,000 rows)
        int totalRows = 10000;
        for (int row = 0; row < totalRows; row++)
        {
            // Populate column A with sequential numbers
            cells[row, 0].PutValue(row + 1);
            // Populate column B with double the value of column A
            cells[row, 1].PutValue((row + 1) * 2);
        }

        // Add a formula that depends on the imported data
        // This formula will sum all values in column A after the import
        cells[totalRows, 0].Formula = $"=SUM(A1:A{totalRows})";

        // Re‑enable automatic calculation (or set to desired mode)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Force calculation of all formulas now that data import is complete
        workbook.CalculateFormula();

        // Save the workbook (lifecycle save)
        workbook.Save("BulkImportResult.xlsx");
    }
}