// Title: How to duplicate consecutive rows with Cells.CopyRows and ensure formulas adjust in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy multiple rows while preserving relative formula references | Using Cells.CopyRows to duplicate rows with formulas in an Excel workbook | Verify that copied rows update B*C formulas after CopyRows in Aspose.Cells | Duplicate data rows and recalculate formulas with Aspose.Cells .NET example | CopyRows method preserving formula calculations when duplicating rows
// Tags: Cells.CopyRows duplicate rows Aspose.Cells | preserve formula references after row copy .NET | copy rows with formulas Excel Xlsx C# | recalculate workbook formulas after row duplication | duplicate consecutive rows Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

namespace DuplicateRowsDemo
{
    // The example creates a workbook, adds header and two data rows with quantity, price, and total formulas, calculates the initial values, then uses Cells.CopyRows to duplicate those rows below the original set, recalculates formulas, prints both original and copied rows, verifies that the total column formulas reference the new rows, and saves the result as an Xlsx file.
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

                // -------------------------------------------------
                // 1. Populate initial data with formulas
                // -------------------------------------------------
                // Header row
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Qty");
                cells["C1"].PutValue("Price");
                cells["D1"].PutValue("Total");

                // First data row (row index 1 -> Excel row 2)
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["C2"].PutValue(2);
                cells["D2"].Formula = "=B2*C2";

                // Second data row (row index 2 -> Excel row 3)
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(5);
                cells["C3"].PutValue(3);
                cells["D3"].Formula = "=B3*C3";

                // Calculate formulas for the original rows
                workbook.CalculateFormula();

                Console.WriteLine("=== Original Data ===");
                PrintRow(cells, 1); // Row 2
                PrintRow(cells, 2); // Row 3

                // -------------------------------------------------
                // 2. Duplicate the two data rows (rows 2 and 3) below them
                // -------------------------------------------------
                // Parameters:
                // sourceRowIndex      = 1 (first data row, zero‑based)
                // destinationRowIndex = 3 (row after the original data rows)
                // totalRows           = 2 (number of rows to copy)
                try
                {
                    // CopyRows requires the source Cells collection as the first argument
                    cells.CopyRows(cells, 1, 3, 2);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error during row copy: " + ex.Message);
                    throw;
                }

                // Re‑calculate to obtain the new values.
                workbook.CalculateFormula();

                Console.WriteLine("\n=== After CopyRows ===");
                // Original rows (unchanged)
                PrintRow(cells, 1);
                PrintRow(cells, 2);
                // Copied rows (rows 4 and 5 in Excel -> indices 3 and 4)
                PrintRow(cells, 3);
                PrintRow(cells, 4);

                // -------------------------------------------------
                // 3. Verify that the formulas have been updated correctly
                // -------------------------------------------------
                Console.WriteLine("\n=== Formula Verification ===");
                Console.WriteLine($"D2 Formula: {cells["D2"].Formula} (expected =B2*C2)");
                Console.WriteLine($"D3 Formula: {cells["D3"].Formula} (expected =B3*C3)");
                Console.WriteLine($"D4 Formula: {cells["D4"].Formula} (expected =B4*C4)");
                Console.WriteLine($"D5 Formula: {cells["D5"].Formula} (expected =B5*C5)");

                // -------------------------------------------------
                // 4. Save the workbook to a file for visual inspection
                // -------------------------------------------------
                string outFile = "DuplicateRowsDemo.xlsx";

                // Ensure the directory exists before saving (if a path is provided)
                string outDir = Path.GetDirectoryName(Path.GetFullPath(outFile));
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"\nWorkbook saved to '{outFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        /// <param name="cells">The Cells collection.</param>
        /// <param name="rowIndex">Zero‑based row index to print.</param>
        private static void PrintRow(Cells cells, int rowIndex)
        {
            string item = cells[rowIndex, 0].StringValue;
            double qty = cells[rowIndex, 1].DoubleValue;
            double price = cells[rowIndex, 2].DoubleValue;
            double total = cells[rowIndex, 3].DoubleValue;

            Console.WriteLine($"Row {rowIndex + 1}: Item={item}, Qty={qty}, Price={price}, Total={total}");
        }
    }
}
