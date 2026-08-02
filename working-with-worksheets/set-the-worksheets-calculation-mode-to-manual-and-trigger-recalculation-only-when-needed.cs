// Title: Aspose.Cells .NET: Set Manual Calculation Mode and Trigger Formula Recalculation
// Description: Demonstrates how to switch a workbook to manual calculation mode, disable automatic formula evaluation on save, modify cell values later, and explicitly invoke CalculateFormula to update dependent formulas using Aspose.Cells for .NET.
// Keywords: Aspose.Cells manual calculation | CalcModeType.Manual | disable automatic formula calculation | CalculateFormula C# | Aspose.Cells performance optimization | manual workbook recalculation | formula settings Aspose.Cells | .NET spreadsheet API
// Common Searches: Aspose.Cells set manual calculation mode | how to prevent formula calculation on save Aspose.Cells | trigger manual recalculation with CalculateFormula | manual vs automatic calculation Aspose.Cells .NET | optimize large workbook performance Aspose.Cells
// Developer Intent: Configure a workbook to use manual calculation and recalculate formulas only when explicitly requested.
// Use Cases: Improve performance when generating large workbooks with thousands of formulas by postponing calculation until all data is populated. | Create a template that users can fill out without triggering any calculations until the file is processed on the server. | Update specific input cells in a saved workbook and refresh dependent results on demand before final export.
// AI Prompts: Show me C# code to set Aspose.Cells workbook to manual calculation mode and manually invoke formula evaluation. | How can I disable automatic formula calculation on save with Aspose.Cells and later recalculate only changed cells? | Provide an example of using CalculateFormula after modifying a cell in a manually calculated workbook.

using System;
using Aspose.Cells;

// Demonstrates how to switch a workbook to manual calculation mode, disable automatic formula evaluation on save, modify cell values later, and explicitly invoke CalculateFormula to update dependent formulas using Aspose.Cells for .NET.
class ManualCalculationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data and formulas
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].Formula = "=A1+A2"; // Sum of A1 and A2

        // Set calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Prevent automatic calculation on save
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Save the workbook (no calculation performed)
        workbook.Save("ManualMode.xlsx");

        // Load the workbook later
        Workbook loadedWb = new Workbook("ManualMode.xlsx");
        Worksheet loadedSheet = loadedWb.Worksheets[0];

        // Change a value that affects the formula
        loadedSheet.Cells["A1"].PutValue(20);

        // Manually trigger calculation only when needed
        loadedWb.CalculateFormula();

        // Save the workbook after manual calculation
        loadedWb.Save("ManualMode_Calculated.xlsx");
    }
}
