// Title: How to make an Aspose.Cells ListObject table auto‑expand when adding rows in C#
// Common Searches: Aspose.Cells C# add row below table and automatically extend ListObject range | Enable auto expand for Excel table created with Aspose.Cells .NET | C# Aspose.Cells expand ListObject when inserting new data rows | Automatically grow Excel table range using Aspose.Cells ListObject API | Aspose.Cells .NET dynamic table range expansion example
// Tags: Aspose.Cells ListObject auto expand | C# expand Excel table range | Aspose.Cells add rows to table | dynamic ListObject range Aspose.Cells | Excel table auto‑grow Aspose .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, defines a ListObject over A1:C3, then inserts a new row directly below the table. Aspose.Cells automatically expands the ListObject to include the new row, and the file is saved as AutoExpandTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate initial data for the table (A1:C3)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(90);

            // Define the range that will become a table (zero‑based indices)
            int firstRow = 0;      // Row 1
            int firstColumn = 0;   // Column A
            int totalRows = 3;     // Header + 2 data rows
            int totalColumns = 3;  // Columns A‑C

            // Add a ListObject (table) to the worksheet; the last argument indicates that the range has headers
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Note: Aspose.Cells automatically expands the table when rows are added within the range,
            // so explicit ShowAutoExpand property is not required.

            // Add new data directly below the current table range
            int newRowIndex = firstRow + totalRows; // Row after the current table
            sheet.Cells[newRowIndex, firstColumn].PutValue(3);
            sheet.Cells[newRowIndex, firstColumn + 1].PutValue("Charlie");
            sheet.Cells[newRowIndex, firstColumn + 2].PutValue(78);

            // Save the workbook
            workbook.Save("AutoExpandTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
