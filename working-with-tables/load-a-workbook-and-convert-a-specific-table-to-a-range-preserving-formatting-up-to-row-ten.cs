using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ConvertTableToRange
{
    static void Main()
    {
        // Load the workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one table (ListObject) on the sheet
        if (sheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No tables found in the worksheet.");
            return;
        }

        // Retrieve the first table
        ListObject table = sheet.ListObjects[0];

        // Define conversion options: convert only up to row 10 (zero‑based index 9)
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 9
        };

        // Convert the table to a normal range while preserving formatting
        table.ConvertToRange(options);

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}