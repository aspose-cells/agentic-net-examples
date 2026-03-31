using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with a header row
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Score");
        cells["A2"].PutValue(1);
        cells["B2"].PutValue("Alice");
        cells["C2"].PutValue(85);
        cells["A3"].PutValue(2);
        cells["B3"].PutValue("Bob");
        cells["C3"].PutValue(92);

        // Define the range for the ListObject (table) using zero‑based indices
        int startRow = 0;      // Row 1
        int startColumn = 0;   // Column A
        int endRow = 2;        // Row 3
        int endColumn = 2;     // Column C
        bool hasHeaders = true;

        // Add the ListObject to the worksheet
        int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, hasHeaders);
        ListObject table = sheet.ListObjects[tableIndex];
        table.DisplayName = "SampleTable";

        // Save the workbook as an XLSX file
        workbook.Save("SampleListObject.xlsx", SaveFormat.Xlsx);
    }
}