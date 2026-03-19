using System;
using Aspose.Cells;

namespace MinMaxIfsDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and populate sample data (including duplicates)
            // -----------------------------------------------------------------
            Workbook wbCreate = new Workbook();
            Worksheet wsCreate = wbCreate.Worksheets[0];
            Cells cellsCreate = wsCreate.Cells;

            // Header
            cellsCreate["A1"].PutValue("Category");
            cellsCreate["B1"].PutValue("Value");

            // Sample data with intentional duplicate rows
            string[] categories = { "Apple", "Banana", "Apple", "Orange", "Banana", "Apple", "Orange", "Apple" };
            double[] values = { 10, 20, 5, 30, 25, 15, 35, 8 };

            for (int i = 0; i < categories.Length; i++)
            {
                cellsCreate[i + 1, 0].PutValue(categories[i]); // Column A
                cellsCreate[i + 1, 1].PutValue(values[i]);    // Column B
            }

            // MINIFS: minimum value where Category = "Apple"
            cellsCreate["D1"].PutValue("MinIfsResult");
            cellsCreate["D2"].Formula = "=MINIFS(B2:B9, A2:A9, \"Apple\")";

            // MAXIFS: maximum value where Category = "Apple"
            cellsCreate["E1"].PutValue("MaxIfsResult");
            cellsCreate["E2"].Formula = "=MAXIFS(B2:B9, A2:A9, \"Apple\")";

            // Save the workbook to disk (using the provided save method)
            string filePath = "MinMaxIfsDemo.xlsx";
            wbCreate.Save(filePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 2. Load the workbook (using the provided load method)
            // -----------------------------------------------------------------
            Workbook wbLoad = new Workbook(filePath);
            Worksheet wsLoad = wbLoad.Worksheets[0];
            Cells cellsLoad = wsLoad.Cells;

            // -----------------------------------------------------------------
            // 3. Remove duplicate rows in the data range (A2:B9)
            //    This ensures that MINIFS and MAXIFS operate on unique records.
            // -----------------------------------------------------------------
            // Determine the actual last row with data
            int lastRow = cellsLoad.MaxDataRow; // zero‑based index
            // Remove duplicates in the range A2:B{lastRow}
            cellsLoad.RemoveDuplicates(1, 0, lastRow, 1, true, new int[] { 0, 1 });

            // -----------------------------------------------------------------
            // 4. Calculate all formulas (including MINIFS and MAXIFS)
            // -----------------------------------------------------------------
            wbLoad.CalculateFormula();

            // -----------------------------------------------------------------
            // 5. Output the evaluated results
            // -----------------------------------------------------------------
            Console.WriteLine("After removing duplicates and calculating formulas:");
            Console.WriteLine($"MINIFS result (minimum Apple value): {cellsLoad["D2"].Value}");
            Console.WriteLine($"MAXIFS result (maximum Apple value): {cellsLoad["E2"].Value}");
        }
    }
}