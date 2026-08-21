// Title: Aspose.Cells .NET – Set Workbook to Manual Calculation Mode and Recalculate On Demand
// Description: Demonstrates how to switch a workbook to manual calculation using CalcModeType.Manual, disable automatic formula evaluation on save, save the file unchanged, then explicitly invoke CalculateFormula and save the updated results. Includes code for populating cells, applying a SUM formula, and managing calculation flow in C#.
// Keywords: Aspose.Cells manual calculation | CalcModeType.Manual C# | disable automatic formula calculation | CalculateFormula example | Aspose.Cells workbook save without calculation | on‑demand formula evaluation | C# Aspose.Cells performance optimization
// Common Searches: Aspose.Cells set manual calculation mode | How to prevent formula calculation on save Aspose.Cells | Trigger CalculateFormula after manual mode .NET | Aspose.Cells disable automatic recalculation | Manual formula evaluation with Aspose.Cells C#
// Developer Intent: Configure a workbook to use manual calculation and run formulas only when explicitly requested.
// Use Cases: Create large spreadsheets where intermediate saves must not incur heavy formula processing, then calculate once before final export. | Generate template files with placeholder formulas that downstream systems will evaluate at a later stage. | Programmatically modify cell values and refresh dependent formulas only when the report is ready for publishing.
// AI Prompts: Show how to set manual calculation mode in Aspose.Cells for .NET and call CalculateFormula later. | Provide a C# example that disables automatic calculation on save, saves the workbook before and after manual recalculation, and explains the behavior. | Explain how to toggle between manual and automatic calculation modes using Aspose.Cells API.

using Aspose.Cells;
using System;

// Demonstrates how to switch a workbook to manual calculation using CalcModeType.Manual, disable automatic formula evaluation on save, save the file unchanged, then explicitly invoke CalculateFormula and save the updated results. Includes code for populating cells, applying a SUM formula, and managing calculation flow in C#.
class ManualCalculationExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells with values and a formula
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Set the calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Prevent automatic calculation on save while in manual mode
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Save the workbook without calculating the formula
        workbook.Save("ManualMode_NoCalculation.xlsx");

        // When recalculation is required, invoke it explicitly
        workbook.CalculateFormula();

        // Save the workbook after manual calculation
        workbook.Save("ManualMode_WithCalculation.xlsx");
    }
}
