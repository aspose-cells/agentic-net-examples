// Title: Copy a worksheet with formulas while preserving calculation mode using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a source sheet with a table of formulas, set the workbook to manual calculation mode, duplicate the sheet to a new worksheet using CopyOptions (ReferToSheetWithSameName), verify that the calculation mode remains unchanged, optionally recalculate formulas, and save the file.
// Keywords: Aspose.Cells copy worksheet | duplicate sheet formulas .NET | preserve calculation mode | CopyOptions ReferToSheetWithSameName | manual calculation mode Aspose.Cells | C# Aspose.Cells example | calculate formulas programmatically | Excel workbook cloning
// Common Searches: Aspose.Cells copy sheet with formulas | keep manual calculation setting after copying worksheet | CopyOptions ReferToSheetWithSameName usage | how to duplicate a worksheet without triggering recalculation | C# copy Excel sheet preserving formula settings
// Developer Intent: Duplicate a worksheet that contains formulas and ensure the workbook's calculation mode stays unchanged.
// Use Cases: Create a template sheet with complex formulas and generate multiple report tabs without automatic recalculation. | Clone a data‑entry worksheet for different departments while maintaining manual calculation for performance. | Programmatically copy sheets and invoke CalculateFormula only when final results are required.
// AI Prompts: Show C# code to copy an Aspose.Cells worksheet with formulas and keep the calculation mode manual. | Explain how CopyOptions.ReferToSheetWithSameName works when duplicating a sheet in Aspose.Cells. | Provide steps to preserve workbook calculation settings during sheet copy and trigger formula calculation on demand.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a source sheet with a table of formulas, set the workbook to manual calculation mode, duplicate the sheet to a new worksheet using CopyOptions (ReferToSheetWithSameName), verify that the calculation mode remains unchanged, optionally recalculate formulas, and save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Rename the default worksheet to "Source"
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate a simple table with formulas in the source sheet
            sourceSheet.Cells["A1"].PutValue("Item");
            sourceSheet.Cells["B1"].PutValue("Quantity");
            sourceSheet.Cells["C1"].PutValue("Price");
            sourceSheet.Cells["D1"].PutValue("Total");

            sourceSheet.Cells["A2"].PutValue("Apple");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["C2"].PutValue(2);
            sourceSheet.Cells["D2"].Formula = "=B2*C2";

            sourceSheet.Cells["A3"].PutValue("Banana");
            sourceSheet.Cells["B3"].PutValue(5);
            sourceSheet.Cells["C3"].PutValue(1.5);
            sourceSheet.Cells["D3"].Formula = "=B3*C3";

            // Set the workbook calculation mode (e.g., Manual)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Add a new blank worksheet that will receive the copied content
            Worksheet copiedSheet = workbook.Worksheets.Add("Copied");

            // Configure copy options to keep references to sheets with the same name
            CopyOptions copyOptions = new CopyOptions
            {
                ReferToSheetWithSameName = true
            };

            // Copy the source sheet (including formulas) into the new sheet using the options
            copiedSheet.Copy(sourceSheet, copyOptions);

            // Ensure the calculation mode remains unchanged after the copy operation
            workbook.Settings.FormulaSettings.CalculationMode = workbook.Settings.FormulaSettings.CalculationMode;

            // Optionally calculate formulas now (useful if the mode is Manual and you need results)
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("CopyTableWithFormulas.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
