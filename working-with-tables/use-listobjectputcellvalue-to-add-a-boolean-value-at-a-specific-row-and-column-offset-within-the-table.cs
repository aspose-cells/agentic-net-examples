// Title: Insert a Boolean flag into a specific row and column of an Aspose.Cells ListObject table with C#
// AI Prompts: Generate C# code that creates a ListObject, then uses ListObject.PutCellValue to set a true value at row offset 1 and column offset 1. | Show how to add a boolean value to the 'Flag' column of an Aspose.Cells table by specifying row and column offsets. | Provide a C# example that defines a table range, creates a ListObject, and inserts a true flag into a cell using PutCellValue.
// Common Searches: Aspose.Cells C# put boolean value in ListObject table cell by offset | How to set true flag in a specific column of an Excel table using Aspose.Cells | C# ListObject.PutCellValue row offset column offset example | Insert boolean into Aspose.Cells table without referencing cell address | Add flag column value to Aspose.Cells ListObject programmatically
// Tags: Aspose.Cells ListObject PutCellValue boolean | C# insert boolean into Excel table cell | row offset column offset Aspose.Cells | populate Flag column ListObject C# | programmatic table data entry Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// Demonstrates creating a workbook, defining a ListObject range, and using ListObject.PutCellValue with row and column offsets to insert a true boolean into the 'Flag' column, then saving the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to define the table range
            sheet.Cells["A1"].PutValue("ID");   // Header for first column
            sheet.Cells["B1"].PutValue("Flag"); // Header for second column
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);

            // Define the range of the ListObject (table) including headers
            int firstRow = 0;          // Row index for A1 (zero‑based)
            int firstColumn = 0;       // Column index for A1
            int totalRows = 4;         // A1:B4 (including header)
            int totalColumns = 2;      // Two columns: A and B

            // Add the ListObject to the worksheet and give it a name
            int listObjectIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[listObjectIndex];
            table.DisplayName = "MyTable";

            // Add a boolean value using PutCellValue
            // rowOffset = 1 -> second data row (row 3 in the sheet, because row 0 is header)
            // columnOffset = 1 -> second column of the table ("Flag")
            int rowOffset = 1;
            int columnOffset = 1;
            table.PutCellValue(rowOffset, columnOffset, true);

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
