// Title: Aspose.Cells for .NET: Set Workbook to Automatic Calculation, Recalculate Formulas, and Verify Results
// Description: Shows how to create a workbook, fill cells A1‑A3, add formulas to B1‑B3, switch FormulaSettings.CalculationMode to Automatic, run Workbook.CalculateFormula, read the computed values, and save the file.
// Keywords: Aspose.Cells | C# automatic calculation mode | Workbook.CalculateFormula | formula evaluation .NET | Excel automatic recalculation | verify calculated cells | set calculation mode | recalculate all formulas | read cell values | Aspose.Cells API
// Common Searches: Aspose.Cells set calculation mode to Automatic | How to recalculate formulas in Aspose.Cells C# | Read formula results after automatic calculation Aspose.Cells | Force full workbook recalculation Aspose.Cells .NET | Verify cell values after Workbook.CalculateFormula
// Developer Intent: Enable automatic formula calculation, trigger a full recalculation, and programmatically confirm the updated cell values.
// Use Cases: Generate server‑side Excel reports where formulas must be evaluated before delivery. | Validate business logic in automated tests by checking calculated results after data changes. | Prepare workbooks for downstream systems that require static values instead of live formulas.
// AI Prompts: Write C# code that switches a workbook to Manual calculation, updates a range, then recalculates only the affected formulas using Aspose.Cells. | Create a unit test that compares expected and actual values after calling Workbook.CalculateFormula on a workbook with dependent formulas. | Explain how to toggle between Automatic and Manual calculation modes in Aspose.Cells while preserving performance for large sheets.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationDemo
{
    // Shows how to create a workbook, fill cells A1‑A3, add formulas to B1‑B3, switch FormulaSettings.CalculationMode to Automatic, run Workbook.CalculateFormula, read the computed values, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(15);

            // Add formulas that depend on the data above
            cells["B1"].Formula = "=A1*2";      // Expected 10
            cells["B2"].Formula = "=A2*2";      // Expected 20
            cells["B3"].Formula = "=SUM(A1:A3)"; // Expected 30

            // Set the calculation mode to Automatic
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Verify and display the calculated values
            Console.WriteLine("Calculated Values after setting Automatic mode:");
            Console.WriteLine($"B1 (A1*2) = {cells["B1"].Value}");
            Console.WriteLine($"B2 (A2*2) = {cells["B2"].Value}");
            Console.WriteLine($"B3 (SUM(A1:A3)) = {cells["B3"].Value}");

            // Optionally save the workbook to inspect the results
            workbook.Save("CalculationModeAutomatic.xlsx");
        }
    }
}
