// Title: Disable PreserveFormatting on a QueryTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# disable PreserveFormatting for Excel query tables | set QueryTable PreserveFormatting to false before refreshing workbook | remove formatting preservation on QueryTable refresh using Aspose.Cells | how to allow default formatting when refreshing a QueryTable in .NET | change QueryTable formatting behavior with Aspose.Cells API
// Tags: Aspose.Cells QueryTable PreserveFormatting property | disable formatting preservation on Excel query table | C# modify QueryTable refresh settings | Excel workbook query table default formatting | Aspose.Cells workbook save after QueryTable update

using Aspose.Cells;
using System;

// Loads an Excel workbook, accesses the first worksheet, checks for a QueryTable, sets its PreserveFormatting property to false so that default formatting is applied on subsequent refreshes, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the query table
        Worksheet sheet = workbook.Worksheets[0]; // adjust index or name as needed

        // Ensure the worksheet has at least one query table
        if (sheet.QueryTables.Count > 0)
        {
            // Get the first query table (or select by name/index as required)
            QueryTable queryTable = sheet.QueryTables[0];

            // Disable PreserveFormatting so that default formatting is applied on refresh
            queryTable.PreserveFormatting = false;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
