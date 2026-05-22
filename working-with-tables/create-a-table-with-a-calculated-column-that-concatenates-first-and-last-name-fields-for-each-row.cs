using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class TableWithCalculatedColumn
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with headers
        sheet.Cells["A1"].PutValue("FirstName");
        sheet.Cells["B1"].PutValue("LastName");
        sheet.Cells["C1"].PutValue("FullName"); // Header for the calculated column

        // Sample rows
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue("Doe");

        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue("Smith");

        sheet.Cells["A4"].PutValue("Bob");
        sheet.Cells["B4"].PutValue("Johnson");

        // Define the range that will become a table (including header row)
        // Rows: 0‑4 (5 rows, 0‑based), Columns: 0‑2 (3 columns)
        int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.DisplayName = "People";
        table.ShowHeaderRow = true;

        // Set the formula for the calculated column using structured references
        // This formula will be applied to each row of the table
        table.ListColumns[2].Formula = "=[@FirstName] & \" \" & [@LastName]";

        // Optionally, show totals row (not required for this task)
        table.ShowTotals = false;

        // Save the workbook
        workbook.Save("TableWithCalculatedColumn.xlsx");
    }
}