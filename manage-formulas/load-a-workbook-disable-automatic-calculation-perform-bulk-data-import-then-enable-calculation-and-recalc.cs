// Title: Disable Automatic Calculation, Bulk Import with ImportObjectArray, Re‑enable and Recalculate Formulas – Aspose.Cells for .NET
// Description: Creates a new Workbook, switches the formula engine to Manual, imports a one‑dimensional object array into the first sheet using ImportObjectArray, restores Automatic mode, forces a full recalculation with CalculateFormula, and saves the result. This approach minimizes intermediate calculations and boosts performance for large data loads.
// Keywords: Aspose.Cells Manual calculation mode | ImportObjectArray C# | CalculateFormula | disable formula recalculation | bulk data import performance | C# workbook save | Excel bulk load Aspose.Cells | formula engine manual | recalculate all formulas | Aspose.Cells settings
// Common Searches: Aspose.Cells turn off formula calculation | ImportObjectArray without triggering recalculation | Set calculation mode to Manual in Aspose.Cells .NET | How to recalculate all formulas after data import | Speed up large Excel write with Aspose.Cells
// Developer Intent: Temporarily suspend formula evaluation, load a large data set efficiently, then reactivate calculation and perform a single comprehensive recalculation before saving the workbook.
// Use Cases: Load a product catalog into a new workbook without intermediate formula updates, then compute totals once. | Generate a financial report where formulas depend on thousands of imported rows, ensuring a single final recalculation. | Populate massive lookup tables for data‑driven dashboards while keeping the calculation engine idle for speed. | Create a template that requires bulk data insertion followed by a full formula refresh prior to distribution.
// AI Prompts: Write C# code that disables automatic calculation, imports a multi‑row object array with ImportObjectArray, re‑enables calculation, runs CalculateFormula, and saves the workbook using Aspose.Cells. | Explain how setting Workbook.Settings.FormulaSettings.CalculationMode to Manual improves performance during large imports and how to safely switch back to Automatic. | Show how to verify that every formula has been recalculated after a bulk import operation in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new Workbook, switches the formula engine to Manual, imports a one‑dimensional object array into the first sheet using ImportObjectArray, restores Automatic mode, forces a full recalculation with CalculateFormula, and saves the result. This approach minimizes intermediate calculations and boosts performance for large data loads.
class BulkImportExample
{
    static void Main()
    {
        // Create a new workbook (uses the Workbook() constructor rule)
        Workbook workbook = new Workbook();

        // Disable automatic calculation by setting the calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Prepare bulk data to import (example data array)
        object[] bulkData = new object[]
        {
            "Product", "Price", "Stock",          // Header row
            "Apple",   1.99,   100,               // First data row
            "Banana",  0.99,   150,               // Second data row
            "Cherry",  2.49,   200                // Third data row
        };

        // Import the object array into the first worksheet starting at cell A1 (row 0, column 0)
        // Imported horizontally (isVertical = false) using the ImportObjectArray method rule
        workbook.Worksheets[0].Cells.ImportObjectArray(bulkData, 0, 0, false);

        // Re‑enable automatic calculation (or set to desired mode)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Recalculate all formulas in the workbook using the CalculateFormula method rule
        workbook.CalculateFormula();

        // Save the workbook to disk using the Save method rule
        workbook.Save("BulkImportResult.xlsx");
    }
}
