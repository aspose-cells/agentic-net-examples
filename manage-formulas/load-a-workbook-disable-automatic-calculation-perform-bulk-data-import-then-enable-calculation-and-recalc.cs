// Title: C# – Load Workbook, Disable Formula Calculation, Bulk Import Data, Re‑enable and Recalculate with Aspose.Cells
// Description: Shows how to open an existing XLSX file using Aspose.Cells for .NET, set the workbook to manual calculation mode to avoid formula evaluation during a bulk import with ImportObjectArray, then restore automatic calculation, force a full recalculation, and save the result.
// Keywords: Aspose.Cells | .NET | C# | Workbook.Load | Manual calculation mode | CalcModeType.Manual | ImportObjectArray | bulk data import | disable formula calculation | recalculate formulas | performance optimization | Excel automation | SaveFormat.Xlsx
// Common Searches: Aspose.Cells disable formula calculation | ImportObjectArray C# example | How to improve import performance Aspose.Cells | Recalculate all formulas after data load Aspose.Cells | Set calculation mode manual Aspose.Cells .NET | Bulk import without triggering formulas Aspose.Cells
// Developer Intent: The developer needs to open a workbook, temporarily suspend formula evaluation, import a large data set efficiently, then reactivate calculation and refresh all formulas before saving.
// Use Cases: Populate a product catalog in a template without repeatedly recalculating totals. | Update financial model inputs programmatically and trigger a single recalculation at the end. | Generate a reporting workbook by bulk‑loading data, then deliver a file with up‑to‑date calculations.
// AI Prompts: Provide C# code that sets Aspose.Cells calculation mode to manual, imports a one‑dimensional object array with ImportObjectArray, then switches back to automatic and runs Workbook.CalculateFormula. | Explain best practices for bulk importing data into an Aspose.Cells workbook while minimizing formula evaluation overhead. | Describe how CalcModeType.Manual and Workbook.CalculateFormula work together after a large data import in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to open an existing XLSX file using Aspose.Cells for .NET, set the workbook to manual calculation mode to avoid formula evaluation during a bulk import with ImportObjectArray, then restore automatic calculation, force a full recalculation, and save the result.
class BulkImportExample
{
    static void Main()
    {
        // Load an existing workbook from file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // uses Workbook(string) constructor

        // Disable automatic calculation by setting manual mode
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare bulk data to import (horizontal layout)
        object[] data = new object[]
        {
            "Product", "Price", "Quantity",   // Header row
            "Apple",   1.20,   10,
            "Banana",  0.80,   20,
            "Cherry",  2.50,   15
        };

        // Import the object array starting at cell A1 (row 0, column 0) horizontally
        sheet.Cells.ImportObjectArray(data, 0, 0, false);

        // Re‑enable automatic calculation
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the modified workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
