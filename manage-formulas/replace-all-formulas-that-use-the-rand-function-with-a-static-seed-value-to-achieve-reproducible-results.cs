// Title: C# – Substitute RAND() with a Fixed Value in Aspose.Cells for Consistent Results
// Description: The example builds a new workbook, inserts volatile RAND() formulas, then walks through each worksheet and cell, swapping the RAND() call for a predetermined constant (e.g., 0.5). After updating the formulas the workbook is recalculated and saved, guaranteeing identical outcomes on every run.
// Keywords: Aspose.Cells static RAND | deterministic formula .NET | replace volatile Excel function | C# workbook formula edit | fixed seed for random Excel
// Common Searches: how to make RAND() deterministic in Aspose.Cells | C# replace random function with constant in Excel | iterate over cells to modify formulas Aspose.Cells | remove volatility from Excel workbook .NET | seeded random value in Aspose.Cells
// Developer Intent: Change all RAND() calls to a predetermined number so the workbook yields the same results each time.
// Use Cases: Prepare financial statements that rely on random numbers, ensuring the published file can be audited with identical values. | Create repeatable test datasets for automated QA by fixing the random component across all sheets. | Stabilize unit‑test scenarios that involve Excel calculations by converting volatile functions to static numbers before execution.
// AI Prompts: Write C# code using Aspose.Cells that scans every worksheet, replaces any RAND() occurrence with a user‑provided constant, recalculates, and saves the workbook. | Show how to iterate through all cells in an Aspose.Cells workbook and substitute volatile functions such as RAND() with fixed values while leaving other formulas untouched. | Explain a method to temporarily replace RAND() with a deterministic value for testing, then restore the original formulas later in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example builds a new workbook, inserts volatile RAND() formulas, then walks through each worksheet and cell, swapping the RAND() call for a predetermined constant (e.g., 0.5). After updating the formulas the workbook is recalculated and saved, guaranteeing identical outcomes on every run.
    public class ReplaceRandWithStaticValue
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // ------------------- Create a workbook (lifecycle rule) -------------------
            Workbook workbook = new Workbook(); // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample formulas that use RAND()
            cells["A1"].Formula = "=RAND() * 100";
            cells["A2"].Formula = "=ROUND(RAND(), 2)";
            cells["A3"].Formula = "=IF(RAND()>0.5, \"Yes\", \"No\")";

            // Calculate original formulas (optional, just to show before replacement)
            workbook.CalculateFormula();

            Console.WriteLine("Before replacement:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 = {cells["A2"].Value}");
            Console.WriteLine($"A3 = {cells["A3"].Value}");

            // ------------------- Replace RAND() with a static value -------------------
            const string randFunction = "RAND()";
            const string staticValue = "0.5"; // deterministic replacement value

            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Cell cell in ws.Cells)
                {
                    if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                        cell.Formula.IndexOf(randFunction, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace all occurrences of RAND() in the formula with the static value
                        string newFormula = cell.Formula.Replace(randFunction, staticValue, StringComparison.OrdinalIgnoreCase);
                        cell.Formula = newFormula;
                    }
                }
            }

            // Recalculate after replacement to get deterministic results
            workbook.CalculateFormula();

            Console.WriteLine("\nAfter replacement:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 = {cells["A2"].Value}");
            Console.WriteLine($"A3 = {cells["A3"].Value}");

            // ------------------- Save the workbook (lifecycle rule) -------------------
            workbook.Save("ReplaceRandWithStaticValue.xlsx"); // save the workbook
        }
    }
}
