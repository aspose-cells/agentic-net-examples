// Title: C# – Update Aspose.Cells ListObject Header Using PutCellValue with Offsets
// Description: Demonstrates how to replace placeholder headers in an Aspose.Cells ListObject by calling ListObject.PutCellValue(rowOffset, columnOffset, value) with a row offset of 0 (header row), then synchronizing column names via UpdateColumnName, and finally saving the workbook.
// Keywords: Aspose.Cells ListObject PutCellValue | C# table header update | Aspose.Cells UpdateColumnName | set ListObject header cell | Aspose.Cells .NET table example
// Common Searches: Aspose.Cells change ListObject header text | PutCellValue row offset 0 header Aspose.Cells | Update table column names after editing header | C# Aspose.Cells ListObject header example | How to rename Aspose.Cells table columns programmatically
// Developer Intent: Replace existing table header values by writing new text to specific header cells with PutCellValue and refresh the ListObject column names.
// Use Cases: Swap placeholder column titles with dynamic names after loading data. | Rename table columns based on user input before exporting a report. | Generate custom report headers that reflect calculated metric names.
// AI Prompts: Show code that updates a ListObject header cell at column index 2 using PutCellValue with row offset 0, then calls UpdateColumnName. | Create a method that accepts a dictionary of column indexes and header strings and updates the ListObject header row via PutCellValue and UpdateColumnName. | Explain how to verify that ListObject column names have been refreshed after modifying header cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to replace placeholder headers in an Aspose.Cells ListObject by calling ListObject.PutCellValue(rowOffset, columnOffset, value) with a row offset of 0 (header row), then synchronizing column names via UpdateColumnName, and finally saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add some initial data with placeholder headers
        cells["A1"].PutValue("OldHeader1");
        cells["B1"].PutValue("OldHeader2");
        cells["A2"].PutValue(10);
        cells["B2"].PutValue(20);

        // Create a ListObject (table) that includes the header row.
        // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
        int tableIndex = sheet.ListObjects.Add(0, 0, 1, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Populate the header cells using row and column offsets.
        // Row offset 0 refers to the header row of the table.
        table.PutCellValue(0, 0, "NewHeader1"); // First header cell (A1)
        table.PutCellValue(0, 1, "NewHeader2"); // Second header cell (B1)

        // Synchronize the ListObject column names with the updated header values.
        table.UpdateColumnName();

        // Save the workbook to a file.
        workbook.Save("ListObjectHeaderUpdate.xlsx", SaveFormat.Xlsx);
    }
}
