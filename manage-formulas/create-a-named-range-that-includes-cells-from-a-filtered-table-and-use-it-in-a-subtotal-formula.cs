// Title: Insert a SUBTOTAL formula that sums visible rows of a filtered ListObject table using Aspose.Cells for C#
// Common Searches: how to add a SUBTOTAL function to a filtered Excel table with Aspose.Cells C# | Aspose.Cells C# calculate sum of visible rows after applying AutoFilter | using ListObject table column reference in SUBTOTAL formula with Aspose.Cells | C# create Excel table, apply AutoFilter and insert SUBTOTAL(9) via Aspose.Cells | Aspose.Cells example for subtotaling filtered data in a worksheet
// Tags: Aspose.Cells SUBTOTAL filtered table | C# ListObject AutoFilter subtotal | Aspose.Cells table column reference formula | Excel workbook calculate visible rows C# | Aspose.Cells create ListObject table

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a ListObject table with sample data, applies an AutoFilter to show only rows where the Category equals "A", inserts a SUBTOTAL(9) formula that sums the visible Sales column values using the table column reference, forces immediate formula calculation, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and name it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate header
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["C1"].PutValue("Sales");

                // Populate sample rows
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(i + 1);                         // ID
                    sheet.Cells[i + 1, 1].PutValue(i % 2 == 0 ? "A" : "B");       // Category
                    sheet.Cells[i + 1, 2].PutValue(100 + i * 10);                 // Sales
                }

                // Add a table (ListObject) covering the data range
                int firstRow = 0;               // zero‑based index
                int firstColumn = 0;
                int totalRows = 11;             // header + 10 data rows
                int totalColumns = 3;

                int tableIndex = sheet.ListObjects.Add(
                    firstRow,
                    firstColumn,
                    firstRow + totalRows - 1,
                    firstColumn + totalColumns - 1,
                    true);

                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable";

                // Apply an AutoFilter on the "Category" column (B) to show only Category "A"
                // Column index is zero‑based, so 1 corresponds to column B
                table.AutoFilter.AddFilter(1, "A");
                // Refresh the filter to apply the criteria
                table.AutoFilter.Refresh();

                // Insert a SUBTOTAL formula that sums the visible (filtered) sales values.
                // Use the table column reference directly without a named range.
                sheet.Cells["E2"].Formula = "=SUBTOTAL(9, SalesTable[Sales])";

                // Calculate formulas immediately
                workbook.CalculateFormula();

                // Define output file path
                string outputPath = "FilteredTableWithSubtotal.xlsx";

                // Ensure the directory exists (handle possible null from GetDirectoryName)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
