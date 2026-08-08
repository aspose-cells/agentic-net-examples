// Title: Aspose.Cells C# – Set Workbook to Automatic Calculation Mode and Verify Real‑Time Formula Updates
// Description: Shows how to enable automatic calculation in an Aspose.Cells workbook, assign values to A1 and B1, create a dependent formula in C1, trigger recalculation, modify a source cell, confirm the updated result, and save the file.
// Keywords: Aspose.Cells | C# | automatic calculation mode | CalcModeType.Automatic | formula recalculation | dependent cells update | CalculateFormula | Excel export | real‑time formula update
// Common Searches: Aspose.Cells set calculation mode to automatic | C# Aspose.Cells recalculate formulas after changing cell value | How to enable auto‑recalc in Aspose.Cells .NET | Aspose.Cells automatic formula update not working | CalculateFormula method Aspose.Cells example
// Developer Intent: Demonstrate configuring a workbook for automatic formula calculation and confirming that dependent cells reflect changes instantly.
// Use Cases: Financial models where input changes instantly recalculate totals. | Dynamic reports that update summary formulas after programmatic data edits. | Excel export routines that ensure all formulas are up‑to‑date before saving.
// AI Prompts: Generate C# code using Aspose.Cells to set workbook calculation mode to Automatic and demonstrate formula update after changing a referenced cell. | Explain why Aspose.Cells requires an explicit CalculateFormula call even in Automatic mode and how to verify the result. | Provide a step‑by‑step guide to enable automatic calculation, modify source cells, and confirm dependent cell values in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to enable automatic calculation in an Aspose.Cells workbook, assign values to A1 and B1, create a dependent formula in C1, trigger recalculation, modify a source cell, confirm the updated result, and save the file.
    public class AutomaticCalculationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set initial values
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);

                // Set a formula that depends on A1 and B1
                cells["C1"].Formula = "=A1+B1";

                // Configure the workbook to use Automatic calculation mode
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // Perform initial calculation
                workbook.CalculateFormula();

                // Verify the result of the dependent cell
                Console.WriteLine("Initial C1 value (should be 30): " + cells["C1"].IntValue);

                // Change a dependent cell value
                cells["A1"].PutValue(40);

                // Since the mode is Automatic, recalculate to reflect the change
                // (Aspose.Cells does not auto‑recalculate, so we invoke it explicitly)
                workbook.CalculateFormula();

                // Verify that the dependent cell updated instantly
                Console.WriteLine("Updated C1 value (should be 60): " + cells["C1"].IntValue);

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("AutomaticCalculationDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AutomaticCalculationDemo.Run();
        }
    }
}
