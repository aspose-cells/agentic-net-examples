// Title: Aspose.Cells C# – Enable Automatic Calculation and Confirm Instant Formula Refresh
// Description: Demonstrates how to set a workbook’s FormulaSettings.CalculationMode to CalcModeType.Automatic, add a dependent formula, modify the source cell, and call CalculateFormula so the linked cell updates immediately. Includes saving the file as an optional step.
// Keywords: Aspose.Cells automatic calculation | CalcModeType.Automatic C# | real‑time formula update | recalculate formulas programmatically | Aspose.Cells workbook settings | C# spreadsheet calculation mode | global spreadsheet automation
// Common Searches: set automatic calculation mode Aspose.Cells .NET | verify dependent cell updates after source change Aspose.Cells | C# example for CalcModeType.Automatic | how to trigger formula recalculation in Aspose.Cells | Aspose.Cells instant formula refresh
// Developer Intent: Configure a workbook to recalculate formulas automatically and ensure that any cell referencing changed data reflects the new value without manual intervention.
// Use Cases: Create a template where users edit input cells and see live formula results. | Generate financial reports that automatically refresh all calculations after data import. | Build an analytics dashboard that maintains up‑to‑date values when underlying data is programmatically altered.
// AI Prompts: Show me C# code to set Aspose.Cells workbook calculation mode to Automatic and verify that a dependent cell updates instantly. | Provide a step‑by‑step example of modifying a source cell and triggering an immediate formula recalculation with Aspose.Cells. | Explain how to configure Aspose.Cells for real‑time formula evaluation and save the workbook after changes.

using System;
using Aspose.Cells;

namespace AsposeCellsAutomaticCalculationDemo
{
    // Demonstrates how to set a workbook’s FormulaSettings.CalculationMode to CalcModeType.Automatic, add a dependent formula, modify the source cell, and call CalculateFormula so the linked cell updates immediately. Includes saving the file as an optional step.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set calculation mode to Automatic
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate initial data
            cells["A1"].PutValue(5);                 // Base value
            cells["B1"].Formula = "=A1*2";           // Dependent formula

            // Perform initial calculation
            workbook.CalculateFormula();

            // Display initial results
            Console.WriteLine("Initial values:");
            Console.WriteLine($"A1 = {cells["A1"].IntValue}");
            Console.WriteLine($"B1 (formula result) = {cells["B1"].IntValue}");

            // Change the base cell value
            cells["A1"].PutValue(10);

            // Recalculate to reflect the change instantly
            workbook.CalculateFormula();

            // Display updated results
            Console.WriteLine("\nAfter updating A1:");
            Console.WriteLine($"A1 = {cells["A1"].IntValue}");
            Console.WriteLine($"B1 (updated formula result) = {cells["B1"].IntValue}");

            // Save the workbook (optional)
            workbook.Save("AutomaticCalculationDemo.xlsx");
        }
    }
}
