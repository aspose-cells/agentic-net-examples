using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Set a specific calculation mode (example: AutomaticExceptTable)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;
        // Preserve the original calculation mode for later use
        CalcModeType originalCalcMode = workbook.Settings.FormulaSettings.CalculationMode;

        // -------------------------------------------------
        // Prepare a source worksheet with a simple table that contains formulas
        // -------------------------------------------------
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Header row
        sourceSheet.Cells["A1"].PutValue("Quantity");
        sourceSheet.Cells["B1"].PutValue("Price");
        sourceSheet.Cells["C1"].PutValue("Total");

        // Data rows
        sourceSheet.Cells["A2"].PutValue(2);
        sourceSheet.Cells["B2"].PutValue(10);
        sourceSheet.Cells["A3"].PutValue(5);
        sourceSheet.Cells["B3"].PutValue(7);

        // Formulas that calculate total per row
        sourceSheet.Cells["C2"].Formula = "=A2*B2";
        sourceSheet.Cells["C3"].Formula = "=A3*B3";

        // -------------------------------------------------
        // Add a new worksheet where the table will be copied
        // -------------------------------------------------
        Worksheet destinationSheet = workbook.Worksheets.Add("CopySheet");

        // -------------------------------------------------
        // Configure copy options:
        // - Keep references to sheets with the same name (important when formulas refer to other sheets)
        // -------------------------------------------------
        CopyOptions copyOptions = new CopyOptions
        {
            ReferToSheetWithSameName = true
        };

        // -------------------------------------------------
        // Copy the entire content of the source worksheet to the destination worksheet
        // -------------------------------------------------
        destinationSheet.Copy(sourceSheet, copyOptions);

        // -------------------------------------------------
        // Ensure the workbook's calculation mode remains unchanged
        // -------------------------------------------------
        workbook.Settings.FormulaSettings.CalculationMode = originalCalcMode;

        // Optional: calculate formulas so that the copied sheet shows results immediately
        workbook.CalculateFormula();

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("TableCopyDemo.xlsx");
    }
}