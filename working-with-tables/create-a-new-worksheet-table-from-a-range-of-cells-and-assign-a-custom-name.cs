// Title: Create a worksheet table from a cell range and assign a custom name with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create table from range A1:C5 and set custom table name | Add ListObject to worksheet using zero‑based indices and assign DisplayName in Aspose.Cells | Apply TableStyleMedium9 to a worksheet table programmatically with Aspose.Cells .NET | Save workbook containing a named table as XLSX using Aspose.Cells C# | Set custom name for ListObject table in Aspose.Cells example code
// Tags: add ListObject from cell range Aspose.Cells | set custom table DisplayName C# | apply TableStyleMedium9 programmatically | define table range using zero‑based indices | save workbook with XLSX format containing table

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a new workbook, fills cells A1:C5 with sample data, defines the range using zero‑based row and column indices, adds a ListObject covering that range, sets its DisplayName to "MyCustomTable" and applies TableStyleMedium9, then saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a friendly name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Fill sample data into range A1:C5
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("Diana");
            sheet.Cells["C5"].PutValue(88);

            // Define the range that will become a table (A1:C5)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 5;
            int totalColumns = 3;

            // Add a table (ListObject) to the worksheet based on the defined range
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a custom name and style to the table
            table.DisplayName = "MyCustomTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
