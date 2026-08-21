// Title: Validate Consistency Between Cell.Calculate and Workbook.Calculate in Aspose.Cells (.NET)
// Description: Shows how to calculate an entire workbook with Workbook.CalculateFormula, capture the results, then recalculate selected cells using Cell.Calculate and compare the values to confirm identical outcomes.
// Keywords: Aspose.Cells | .NET | Cell.Calculate | Workbook.CalculateFormula | formula evaluation | calculation consistency | CalculationOptions | unit test | compare results
// Common Searches: Aspose.Cells Cell.Calculate vs Workbook.Calculate | compare formula results Aspose.Cells | does Cell.Calculate give same result as Workbook.Calculate | Aspose.Cells calculate single cell | verify formula calculation consistency Aspose.Cells
// Developer Intent: Confirm that calculating formulas per cell yields the same values as a full‑workbook calculation.
// Use Cases: Create automated unit tests that assert equality between Cell.Calculate and Workbook.Calculate for a set of formula cells. | Debug discrepancies when applying CalculationOptions to individual cells versus the whole workbook. | Generate on‑demand cell values while trusting the results of a prior workbook‑wide calculation. | Benchmark performance differences between cell‑level and workbook‑level calculations.
// AI Prompts: Write a C# NUnit test that runs Workbook.CalculateFormula, then iterates over a list of cell addresses calling Cell.Calculate and asserts equality with a numeric tolerance. | Provide a reusable method that accepts a worksheet, a collection of cell addresses, and CalculationOptions, and verifies that Cell.Calculate matches previously stored Workbook.Calculate results. | Explain how CalculationOptions affect both Cell.Calculate and Workbook.Calculate and outline best practices to keep their outputs consistent.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsComparisonDemo
{
    // Shows how to calculate an entire workbook with Workbook.CalculateFormula, capture the results, then recalculate selected cells using Cell.Calculate and compare the values to confirm identical outcomes.
    class Program
    {
        static void Main()
        {
            // Prepare a list of cell addresses that contain formulas
            string[] formulaCells = { "B1", "C1", "D2", "E3" };

            // ---------- First workbook: calculate whole workbook ----------
            Workbook wbWhole = new Workbook();                     // create workbook
            Worksheet wsWhole = wbWhole.Worksheets[0];
            SetupWorksheet(wsWhole);                              // add data & formulas

            // Calculate all formulas at once
            CalculationOptions options = new CalculationOptions();
            wbWhole.CalculateFormula(options);

            // Store calculated results for later comparison
            var expectedValues = new Dictionary<string, object>();
            foreach (string addr in formulaCells)
            {
                expectedValues[addr] = wsWhole.Cells[addr].Value;
            }

            // ---------- Second workbook: calculate each cell individually ----------
            Workbook wbCell = new Workbook();                      // create another workbook
            Worksheet wsCell = wbCell.Worksheets[0];
            SetupWorksheet(wsCell);                               // same data & formulas (no prior calculation)

            // Compare each cell's individual calculation with the whole‑workbook result
            foreach (string addr in formulaCells)
            {
                Cell cell = wsCell.Cells[addr];
                // Calculate this single cell
                cell.Calculate(options);

                object individualResult = cell.Value;
                object wholeResult = expectedValues[addr];

                bool isEqual = AreValuesEqual(individualResult, wholeResult);
                Console.WriteLine($"Cell {addr}: Individual={individualResult} | Whole={wholeResult} | Consistent={isEqual}");
            }
        }

        // Helper to populate cells with sample data and formulas
        private static void SetupWorksheet(Worksheet sheet)
        {
            Cells cells = sheet.Cells;

            // Input values
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(15);
            cells["B2"].PutValue(2);
            cells["C3"].PutValue(3);

            // Formulas to be tested
            cells["B1"].Formula = "=A1*2";                         // simple multiplication
            cells["C1"].Formula = "=SUM(A1:A3)";                   // sum range
            cells["D2"].Formula = "=AVERAGE(A1:A3)";               // average
            cells["E3"].Formula = "=IF(A1>3,\"High\",\"Low\")";   // conditional text
        }

        // Simple equality check that handles numeric tolerance and nulls
        private static bool AreValuesEqual(object val1, object val2)
        {
            if (val1 == null && val2 == null) return true;
            if (val1 == null || val2 == null) return false;

            // Numeric comparison with tolerance
            if (val1 is double d1 && val2 is double d2)
                return Math.Abs(d1 - d2) < 1e-9;

            // General object.Equals fallback
            return val1.Equals(val2);
        }
    }
}
