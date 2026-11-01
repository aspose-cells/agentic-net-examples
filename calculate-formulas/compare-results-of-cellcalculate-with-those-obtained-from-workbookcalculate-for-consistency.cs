// Title: Validate consistency between Cell.Calculate and Workbook.CalculateFormula results in Aspose.Cells C#
// Common Searches: how to ensure Cell.Calculate returns same value as Workbook.CalculateFormula in Aspose.Cells C# | Aspose.Cells C# compare per‑cell calculation with full workbook calculation | using CalculationOptions with Cell.Calculate to match Workbook.CalculateFormula output | verify formula evaluation consistency after changing input cells in Aspose.Cells | C# example for comparing Cell.Calculate and Workbook.CalculateFormula results
// Tags: Aspose.Cells cell calculate consistency | Aspose.Cells workbook calculateformula comparison | C# calculationoptions with cell.calculate | per‑cell formula evaluation Aspose.Cells | formula result verification Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsCalculateComparison
{
    // The example creates a workbook, assigns numeric values and formulas, calculates the sheet using Workbook.CalculateFormula, then recalculates each cell individually with Cell.Calculate (using CalculationOptions). It compares the resulting values for consistency, modifies the inputs, repeats the calculations, and prints whether the per‑cell and full‑workbook evaluations remain identical.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set input values
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);

                // Set formulas
                cells["C1"].Formula = "A1+B1";
                cells["D1"].Formula = "A1*B1";

                // ---- Calculate using workbook.CalculateFormula() ----
                workbook.CalculateFormula();

                // Retrieve calculated values after workbook.CalculateFormula()
                object workbookCalcC1 = cells["C1"].Value;
                object workbookCalcD1 = cells["D1"].Value;

                // ---- Calculate using individual Cell.Calculate() ----
                // Cell.Calculate requires a CalculationOptions instance
                CalculationOptions options = new CalculationOptions();

                // Calculate each cell and then read its Value property
                cells["C1"].Calculate(options);
                object cellCalcC1 = cells["C1"].Value;

                cells["D1"].Calculate(options);
                object cellCalcD1 = cells["D1"].Value;

                // Compare the results
                bool isC1Consistent = Equals(workbookCalcC1, cellCalcC1);
                bool isD1Consistent = Equals(workbookCalcD1, cellCalcD1);

                // Output the comparison results
                Console.WriteLine($"C1 - Workbook.CalculateFormula: {workbookCalcC1}, Cell.Calculate: {cellCalcC1}, Consistent: {isC1Consistent}");
                Console.WriteLine($"D1 - Workbook.CalculateFormula: {workbookCalcD1}, Cell.Calculate: {cellCalcD1}, Consistent: {isD1Consistent}");

                // Optional: modify input values and repeat the comparison
                cells["A1"].PutValue(5);
                cells["B1"].PutValue(3);

                // Recalculate using workbook.CalculateFormula()
                workbook.CalculateFormula();

                // Retrieve new values
                workbookCalcC1 = cells["C1"].Value;
                workbookCalcD1 = cells["D1"].Value;

                // Recalculate using Cell.Calculate()
                cells["C1"].Calculate(options);
                cellCalcC1 = cells["C1"].Value;

                cells["D1"].Calculate(options);
                cellCalcD1 = cells["D1"].Value;

                // Compare again
                isC1Consistent = Equals(workbookCalcC1, cellCalcC1);
                isD1Consistent = Equals(workbookCalcD1, cellCalcD1);

                Console.WriteLine("After changing inputs:");
                Console.WriteLine($"C1 - Workbook.CalculateFormula: {workbookCalcC1}, Cell.Calculate: {cellCalcC1}, Consistent: {isC1Consistent}");
                Console.WriteLine($"D1 - Workbook.CalculateFormula: {workbookCalcD1}, Cell.Calculate: {cellCalcD1}, Consistent: {isD1Consistent}");
            }
            catch (Exception ex)
            {
                // Runtime safety: log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
