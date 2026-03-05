using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class CreateTableExample
{
    static void Main()
    {
        // Load an existing XLSX file (or create a new one if the file does not exist)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is data to form a table (headers + a few rows)
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");

        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["C2"].PutValue(85);

        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["C3"].PutValue(92);

        // Define the range for the table (including header row)
        int firstRow = 0;          // zero‑based index for row 1 (A1)
        int firstColumn = 0;       // zero‑based index for column A
        int totalRows = 3;         // header + 2 data rows
        int totalColumns = 3;      // columns A, B, C

        // Add a ListObject (Excel table) to the worksheet
        int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
        ListObject table = sheet.ListObjects[listObjectIndex];

        // Set a friendly name and apply a built‑in table style
        table.DisplayName = "SampleTable";
        table.TableStyleType = TableStyleType.TableStyleMedium9;

        // Save the workbook with the newly created table
        workbook.Save("output.xlsx");
    }
}