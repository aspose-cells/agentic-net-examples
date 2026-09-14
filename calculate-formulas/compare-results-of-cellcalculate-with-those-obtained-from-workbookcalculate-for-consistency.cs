// Title: Compare results of Cell.Calculate and Workbook.CalculateFormula for formula consistency in Aspose.Cells C#
// AI Prompts: Write C# code that loops through every formula cell in a worksheet, calls Cell.Calculate on each, and logs any mismatches against the values produced by Workbook.CalculateFormula. | Show how to clone a workbook, run Workbook.CalculateFormula on the clone, then validate that each Cell.Calculate call on the original workbook returns the same result.
// Common Searches: Aspose.Cells C# verify that Cell.Calculate matches Workbook.CalculateFormula output | compare individual cell calculation with full workbook calculation using Aspose.Cells | C# Aspose.Cells formula consistency check between Cell.Calculate and Workbook.CalculateFormula | sample code to iterate over formula cells and compare calculation results in Aspose.Cells | difference in results when using Cell.Calculate versus Workbook.CalculateFormula in Aspose.Cells .NET
// Tags: cell.Calculate formula evaluation consistency | Workbook.CalculateFormula full workbook calculation | Aspose.Cells compare individual vs full calculation | C# iterate formula cells Aspose.Cells | validate formula results Aspose.Cells .NET | clone workbook for calculation comparison Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaComparison
{
    // The example creates a workbook with sample data and formulas, clones it, runs a full workbook calculation on the clone using Workbook.CalculateFormula, then individually calculates each formula cell in the original workbook with Cell.Calculate, and finally compares the values cell‑by‑cell to confirm they are identical.
    class Program
    {
        static void Main()
        {
            // Create a workbook and add some sample data and formulas
            Workbook wbIndividual = new Workbook();
            Worksheet ws = wbIndividual.Worksheets[0];
            Cells cells = ws.Cells;

            // Sample values
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(15);

            // Formulas to test
            cells["B1"].Formula = "=A1*2";
            cells["B2"].Formula = "=SUM(A1:A3)";
            cells["B3"].Formula = "=AVERAGE(A1:A3)";
            cells["C1"].Formula = "=IF(A1>3,\"High\",\"Low\")";

            // ------------------------------------------------------------
            // Clone the workbook to have an identical copy for full workbook calculation
            // ------------------------------------------------------------
            MemoryStream ms = new MemoryStream();
            wbIndividual.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0;
            Workbook wbFull = new Workbook(ms);

            // ------------------------------------------------------------
            // Perform full workbook calculation on the cloned workbook
            // ------------------------------------------------------------
            wbFull.CalculateFormula();

            // ------------------------------------------------------------
            // Perform individual cell calculations on the original workbook
            // ------------------------------------------------------------
            CalculationOptions calcOptions = new CalculationOptions(); // default options
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsFormula)
                    {
                        // Calculate this cell only
                        cell.Calculate(calcOptions);
                    }
                }
            }

            // ------------------------------------------------------------
            // Compare results cell by cell
            // ------------------------------------------------------------
            Console.WriteLine("Comparison of individual cell.Calculate vs workbook.CalculateFormula:");
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cellInd = cells[row, col];
                    Cell cellFull = wbFull.Worksheets[0].Cells[row, col];

                    if (cellInd.IsFormula)
                    {
                        object valInd = cellInd.Value;
                        object valFull = cellFull.Value;

                        bool equal = ObjectEquals(valInd, valFull);
                        string address = cellInd.Name;

                        Console.WriteLine($"{address}: Individual={valInd} | Full={valFull} | Consistent={equal}");
                    }
                }
            }
        }

        // Helper method to compare two cell values considering possible types
        private static bool ObjectEquals(object a, object b)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;

            // Compare numeric values with tolerance
            if (a is double da && b is double db)
                return Math.Abs(da - db) < 1e-9;

            // Compare strings (case‑sensitive)
            if (a is string sa && b is string sb)
                return sa == sb;

            // Compare booleans
            if (a is bool ba && b is bool bb)
                return ba == bb;

            // Fallback to default equality
            return a.Equals(b);
        }
    }
}
