// Title: Changing a ListObject table header cell with PutCellValue row/column offsets in Aspose.Cells for .NET (C#)
// AI Prompts: Use ListObject.PutCellValue to set the header of the second column in an Aspose.Cells table by specifying row offset 0 and column offset 1. | Create a worksheet, add a ListObject, and modify its header row using row and column offsets in C# with Aspose.Cells. | Replace an existing table column header in an Excel file by calling ListObject.PutCellValue with offset parameters in Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells how to update table column header using PutCellValue offsets | c# set ListObject header cell by row and column offset Aspose.Cells example | changing Excel table header programmatically with Aspose.Cells ListObject PutCellValue | Aspose.Cells PutCellValue row offset column offset for table header in .NET | modify ListObject column names after creation using C# Aspose.Cells
// Tags: ListObject.PutCellValue header offset | Aspose.Cells modify table column header | C# Aspose.Cells table header update | Excel ListObject header cell change | Aspose.Cells row column offset example

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating a workbook, adding a ListObject table, and using ListObject.PutCellValue with row and column offsets to change the second column header to "Surname" before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some initial data to define the table range
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Jane");

            // Define the range for the ListObject (table) – includes header row
            int firstRow = 0;          // zero‑based index for row 1 (A1)
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // rows A1:B3
            int totalColumns = 2;      // columns A and B

            // Add the ListObject to the worksheet
            int listObjectIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[listObjectIndex];
            table.DisplayName = "MyTable";

            // Use PutCellValue with row and column offsets to set a header cell.
            // Row offset 0 = header row, column offset 1 = second column header.
            table.PutCellValue(0, 1, "Surname");

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
