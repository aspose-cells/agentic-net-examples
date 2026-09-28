// Title: Enable Automatic Calculation Mode in Aspose.Cells for .NET and Verify Real-Time Dependent Cell Updates
// AI Prompts: Set workbook.Settings.FormulaSettings.CalculationMode to CalcModeType.Automatic, assign a formula to a cell, modify the source cell, call workbook.CalculateFormula, and output the refreshed values. | Persist the automatic calculation setting by saving the workbook after updating cells and invoking CalculateFormula in a C# application.
// Common Searches: Aspose.Cells .NET how to set automatic formula calculation and refresh dependent cells | C# example for enabling CalcModeType.Automatic and forcing recalculation | verify that a formula cell updates immediately after changing a source cell in Aspose.Cells | save workbook with automatic calculation mode enabled using Aspose.Cells for .NET
// Tags: automatic formula calculation setting Aspose.Cells | CalcModeType.Automatic usage C# | manual workbook.CalculateFormula call | dependent cell update verification .NET | persist calculation settings Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsAutomaticCalculationDemo
{
    // Shows how to enable Aspose.Cells' automatic calculation mode in C#, assign a formula, change the source value, invoke CalculateFormula to refresh dependent cells, and save the workbook so the setting persists.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set initial values
            cells["A1"].PutValue(10);   // Source cell
            cells["B1"].Formula = "=A1*2"; // Dependent cell

            // Configure calculation mode to Automatic
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Perform initial calculation
            workbook.CalculateFormula();

            // Verify that dependent cell reflects the source value
            Console.WriteLine("Initial calculation:");
            Console.WriteLine($"A1 = {cells["A1"].IntValue}");
            Console.WriteLine($"B1 (A1*2) = {cells["B1"].IntValue}");

            // Change the source cell value
            cells["A1"].PutValue(20);

            // Recalculate formulas (Aspose does not auto‑recalculate, so we invoke it)
            workbook.CalculateFormula();

            // Verify that dependent cell updates instantly after recalculation
            Console.WriteLine("\nAfter updating A1:");
            Console.WriteLine($"A1 = {cells["A1"].IntValue}");
            Console.WriteLine($"B1 (A1*2) = {cells["B1"].IntValue}");

            // Save the workbook (optional, demonstrates that settings are persisted)
            workbook.Save("AutomaticCalculationDemo.xlsx");
        }
    }
}
