// Title: Show Table Header Row and Lock Header Cells with Aspose.Cells for .NET
// Description: Demonstrates how to add a ListObject to a worksheet, enable its header row, lock the header cells, protect the sheet, and save the workbook as an Excel file using Aspose.Cells in C#.
// Keywords: Aspose.Cells C# | ListObject ShowHeaderRow | lock table header cells | worksheet protection Aspose.Cells | read‑only Excel header | .NET Excel table header | protect header row programmatically
// Common Searches: Aspose.Cells lock table header row | C# show header row ListObject | protect worksheet header cells Aspose | make Excel table header read‑only with Aspose.Cells | how to lock header cells in Aspose.Cells .NET
// Developer Intent: Enable the table header row and make its cells read‑only.
// Use Cases: Create a product catalog where column titles cannot be altered by end users. | Generate compliance‑driven reports that require a fixed header while allowing data entry. | Export data to Excel from an application and ensure the table headings remain unchanged.
// AI Prompts: Provide C# code using Aspose.Cells to add a ListObject, display its header row, lock the header cells, and protect the worksheet. | Show an example that locks only the header row of an Excel table while keeping data rows editable with Aspose.Cells for .NET. | Explain the steps to configure cell style locking and worksheet protection to make a table header read‑only in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableHeaderLockDemo
{
    // Demonstrates how to add a ListObject to a worksheet, enable its header row, lock the header cells, protect the sheet, and save the workbook as an Excel file using Aspose.Cells in C#.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data with a header row
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(2.5);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(1.8);

                // Add a table (ListObject) that includes the header row
                // Parameters: first row, first column, last row, last column, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Ensure the header row is visible
                table.ShowHeaderRow = true;

                // Lock the header cells to prevent editing
                int headerRow = table.StartRow;
                for (int col = table.StartColumn; col <= table.EndColumn; col++)
                {
                    Cell headerCell = worksheet.Cells[headerRow, col];
                    Style style = headerCell.GetStyle();
                    style.IsLocked = true;
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
