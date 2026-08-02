// Title: C# Aspose.Cells Example – Automatic Calculation Mode with Forced Full Recalculation
// Description: Shows how to configure an Aspose.Cells Workbook for automatic formula calculation (CalcModeType.Automatic), enable ForceFullCalculation, add data and a SUM formula, trigger a full recalculation after cell updates, and save the file.
// Keywords: Aspose.Cells automatic calculation | ForceFullCalculation C# | CalcModeType.Automatic | Aspose.Cells FormulaSettings | Workbook.CalculateFormula .NET | Excel formula recalculation C# | Aspose.Cells sample code | full workbook recalculation | C# Excel automation | Aspose.Cells API example
// Common Searches: Aspose.Cells enable automatic calculation C# | ForceFullCalculation Aspose.Cells .NET example | CalcModeType.Automatic usage | Recalculate formulas after cell change Aspose.Cells | Workbook.CalculateFormula C# tutorial | C# code to force full recalculation in Excel | Aspose.Cells automatic vs manual calculation
// Developer Intent: Configure a workbook to recalculate formulas automatically and guarantee a complete recalculation whenever data changes.
// Use Cases: Automatically update totals, averages, or other aggregates when source cells are edited. | Ensure consistent results in complex financial models after bulk data imports. | Maintain accurate calculations in generated reports by forcing a full recompute before saving. | Debug spreadsheet logic by forcing a fresh evaluation of all formulas after each change.
// AI Prompts: Generate C# code that sets Aspose.Cells Workbook to Automatic calculation mode and forces a full recalculation after each data modification. | Provide a step‑by‑step explanation of FormulaSettings.ForceFullCalculation and when to use it in Aspose.Cells. | Create a C# snippet that switches between automatic and manual calculation modes in Aspose.Cells and demonstrates Workbook.CalculateFormula usage. | Compare the performance impact of enabling ForceFullCalculation versus default incremental recalculation in large workbooks.

using System;
using Aspose.Cells;

namespace AsposeCellsAutomaticCalculationDemo
{
    // Shows how to configure an Aspose.Cells Workbook for automatic formula calculation (CalcModeType.Automatic), enable ForceFullCalculation, add data and a SUM formula, trigger a full recalculation after cell updates, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Enable automatic calculation mode
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Force a full recalculation each time a calculation is triggered
            workbook.Settings.FormulaSettings.ForceFullCalculation = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add initial data and a formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=SUM(A1:A2)";

            // Perform the first calculation
            workbook.CalculateFormula();

            Console.WriteLine($"Initial result (A3): {cells["A3"].Value}");

            // Change data to trigger recalculation
            cells["A1"].PutValue(30);
            cells["A2"].PutValue(40);

            // Because ForceFullCalculation is true, a full recalculation will be performed
            workbook.CalculateFormula();

            Console.WriteLine($"After data change (A3): {cells["A3"].Value}");

            // Save the workbook (lifecycle rule)
            workbook.Save("AutomaticCalculationDemo.xlsx");
        }
    }
}
