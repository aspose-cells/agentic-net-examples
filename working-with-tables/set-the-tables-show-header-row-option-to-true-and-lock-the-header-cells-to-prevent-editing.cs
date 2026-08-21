// Title: Lock Table Header Row and Show Headers with Aspose.Cells in C#
// Description: Demonstrates how to add a ListObject table to a new workbook, enable the header row, lock each header cell, protect the worksheet, and save the file as TableHeaderLocked.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | ListObject header lock | ShowHeaderRow | protect worksheet | lock table header cells | Excel table read‑only header | Aspose.Cells example | C# Excel automation | GitHub Aspose.Cells sample | Aspose.Cells table protection
// Common Searches: Aspose.Cells lock table header C# | ShowHeaderRow true Aspose.Cells | protect worksheet after locking cells Aspose.Cells | read‑only header row Aspose.Cells ListObject | C# code to lock Excel table header with Aspose | Aspose.Cells example for header row protection
// Developer Intent: Display the table header row and make the header cells read‑only by locking them and protecting the worksheet.
// Use Cases: Create a template where column titles stay immutable while users fill data rows. | Distribute a report that guarantees consistent header formatting across all recipients. | Build a data‑entry workbook that prevents accidental changes to header labels.
// AI Prompts: Generate C# code with Aspose.Cells that adds a ListObject, sets ShowHeaderRow to true, locks the header cells, and protects the worksheet. | Show an Aspose.Cells example that makes only the table header row read‑only while allowing edits in the data rows. | Explain step‑by‑step how to lock a table's header row and protect the sheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableHeaderLockDemo
{
    // Demonstrates how to add a ListObject table to a new workbook, enable the header row, lock each header cell, protect the worksheet, and save the file as TableHeaderLocked.xlsx using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (including header row)
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(1.20);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(0.80);

                // Add a table (ListObject) that includes the header row
                // Parameters: first row, first column, last row, last column, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Ensure the header row is visible
                table.ShowHeaderRow = true;

                // Determine the range of header cells
                int headerRow = table.StartRow;
                int firstCol = table.StartColumn;
                int lastCol = firstCol + table.ListColumns.Count - 1; // use ListColumns.Count

                // Lock each header cell
                for (int col = firstCol; col <= lastCol; col++)
                {
                    Cell headerCell = worksheet.Cells[headerRow, col];
                    Style style = headerCell.GetStyle();
                    style.IsLocked = true; // Mark cell as locked
                    headerCell.SetStyle(style);
                }

                // Protect the worksheet so that locked cells cannot be edited
                worksheet.Protect(ProtectionType.All);

                // Save the workbook
                workbook.Save("TableHeaderLocked.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
