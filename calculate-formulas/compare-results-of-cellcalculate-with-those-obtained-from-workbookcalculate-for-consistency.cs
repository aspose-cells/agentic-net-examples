// Title: Aspose.Cells .NET – Verify Consistency Between cell.Calculate and workbook.CalculateFormula
// Description: Sample C# program that fills a workbook with data and formulas, runs workbook.CalculateFormula, stores each formula result, then re‑evaluates the same cells with cell.Calculate (using default CalculationOptions) and reports any mismatches before saving the file.
// Keywords: Aspose.Cells cell.Calculate | Aspose.Cells workbook.CalculateFormula | formula calculation consistency .NET | Aspose.Cells unit test example | C# Aspose.Cells calculation comparison | GitHub Aspose.Cells formula validation
// Common Searches: cell.Calculate vs workbook.CalculateFormula Aspose.Cells | Aspose.Cells compare individual cell calculation with full workbook calculation | how to validate formula results in Aspose.Cells .NET | Aspose.Cells C# example for formula consistency check
// Developer Intent: Ensure that invoking Cell.Calculate on each formula cell yields the same values as a bulk workbook.CalculateFormula execution.
// Use Cases: Create automated tests that detect discrepancies between per‑cell and whole‑workbook calculations. | Debug workbooks that contain volatile functions or external references by pinpointing cells with divergent results. | Validate custom CalculationOptions settings before deploying large‑scale spreadsheet processing.
// AI Prompts: Write C# code that logs cells where cell.Calculate differs from workbook.CalculateFormula, including expected and actual values with precision handling. | Show how to set up CalculationOptions for iterative functions and then compare the outcomes of cell.Calculate and workbook.CalculateFormula. | Provide a method that returns a collection of mismatched cells together with their formulas and both result sets.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaComparison
{
    // Sample C# program that fills a workbook with data and formulas, runs workbook.CalculateFormula, stores each formula result, then re‑evaluates the same cells with cell.Calculate (using default CalculationOptions) and reports any mismatches before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(15);

            // Set formulas that reference the data above
            cells["B1"].Formula = "=A1*2";
            cells["B2"].Formula = "=A2+5";
            cells["B3"].Formula = "=SUM(A1:A3)";
            cells["C1"].Formula = "=IF(A1>3,\"High\",\"Low\")";
            cells["C2"].Formula = "=VLOOKUP(10,A1:B3,2,FALSE)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Store the calculated results of each formula cell
            var formulaResults = new Dictionary<string, object>();
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula)
                {
                    // Use the cell's address as the key (e.g., "B1")
                    formulaResults[cell.Name] = cell.Value;
                }
            }

            // Prepare calculation options (default options)
            CalculationOptions options = new CalculationOptions();

            // Re‑calculate each formula cell individually and compare the results
            bool allMatch = true;
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula)
                {
                    // Re‑calculate the cell
                    cell.Calculate(options);

                    // Retrieve the previously stored result
                    object expected = formulaResults[cell.Name];
                    object actual = cell.Value;

                    // Compare using Equals (handles numbers, strings, booleans, etc.)
                    if (!object.Equals(expected, actual))
                    {
                        allMatch = false;
                        Console.WriteLine($"Mismatch in cell {cell.Name}: expected {expected}, got {actual}");
                    }
                }
            }

            if (allMatch)
            {
                Console.WriteLine("All cell.Calculate results match workbook.CalculateFormula results.");
            }

            // Optional: save the workbook to verify lifecycle handling (create/save)
            workbook.Save("FormulaComparisonResult.xlsx");
        }
    }
}
