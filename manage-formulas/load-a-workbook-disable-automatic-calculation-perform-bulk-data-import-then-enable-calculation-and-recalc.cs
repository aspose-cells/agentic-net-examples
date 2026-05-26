using System;
using Aspose.Cells;

namespace BulkImportAndRecalcDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Disable automatic calculation while importing data
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            workbook.Settings.FormulaSettings.CalculateOnOpen = false;
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare bulk data to import (horizontal layout)
            object[] bulkData = new object[]
            {
                "Product", "Quantity", "Price",          // Header row
                "Apple",   10,       1.5,               // Row 2
                "Banana",  20,       0.8,               // Row 3
                "Cherry",  15,       2.2                // Row 4
            };

            // Import the array starting at cell A1 (row 0, column 0) horizontally
            sheet.Cells.ImportObjectArray(bulkData, 0, 0, false);

            // Re‑enable automatic calculation
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
            workbook.Settings.FormulaSettings.CalculateOnOpen = true;
            workbook.Settings.FormulaSettings.CalculateOnSave = true;

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}