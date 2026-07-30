// Title: C# – Update Aspose.Cells ListObject Header Using PutCellValue (row/column offsets)
// Description: Creates a workbook, adds a ListObject with headers, then replaces the header cells by calling ListObject.PutCellValue(rowOffset, columnOffset, newValue) where rowOffset = 0 targets the header row. Saves the file as an XLSX.
// Keywords: Aspose.Cells | ListObject | PutCellValue | C# | Excel table header | row offset | column offset | update table column name | modify Excel header programmatically
// Common Searches: Aspose.Cells change ListObject header cell | PutCellValue row offset 0 header Aspose | C# update Excel table column header Aspose.Cells | How to rename ListObject columns with Aspose.Cells | Set new header names for a table in .NET
// Developer Intent: Replace existing header values of a ListObject table by using ListObject.PutCellValue with specific row and column offsets.
// Use Cases: Rename table columns after importing data from external sources. | Apply user‑defined or localized header titles during report generation. | Standardize column headings across multiple worksheets in an automated export.
// AI Prompts: Generate C# code that loops through a dictionary to rename all ListObject headers using PutCellValue. | Show how to read current ListObject header names and replace them with values from a JSON file. | Explain the meaning of row and column offsets in ListObject.PutCellValue for updating header rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectHeaderDemo
{
    // Creates a workbook, adds a ListObject with headers, then replaces the header cells by calling ListObject.PutCellValue(rowOffset, columnOffset, newValue) where rowOffset = 0 targets the header row. Saves the file as an XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data including initial headers
            worksheet.Cells["A1"].PutValue("OldHeader1");
            worksheet.Cells["B1"].PutValue("OldHeader2");
            worksheet.Cells["A2"].PutValue(10);
            worksheet.Cells["B2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);
            worksheet.Cells["B3"].PutValue(40);

            // Add a ListObject (table) that includes the header row
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Use PutCellValue with row and column offsets to change the header cell
            // Row offset 0 = header row, column offset 0 = first column
            table.PutCellValue(0, 0, "NewHeader1");
            // Optionally change the second header as well
            table.PutCellValue(0, 1, "NewHeader2");

            // Save the workbook
            workbook.Save("ListObjectHeaderUpdated.xlsx", SaveFormat.Xlsx);
        }
    }
}
