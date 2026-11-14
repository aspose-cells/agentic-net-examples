// Title: How to set incremental row heights for multiple rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells set row height in a for loop | increase each row height by 2 points using Aspose.Cells | programmatically assign different heights to rows 0-9 in Excel with Aspose.Cells | example of SetRowHeight with incremental values in C# | automate dynamic row height changes in an Excel file using Aspose.Cells .NET
// Tags: Aspose.Cells SetRowHeight loop | C# incremental Excel row height | Aspose.Cells dynamic row height assignment | Excel row height automation .NET | SetRowHeight multiple rows Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, and uses a for‑loop to set the height of rows 0 through 9. Each row starts at 15 points and grows by 2 points per row via sheet.Cells.SetRowHeight(i, height), then saves the workbook as RowHeights.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set row heights for rows 0 through 9 with incremental values
        for (int i = 0; i < 10; i++)
        {
            // Example: start at 15 points and increase by 2 points each row
            double height = 15 + i * 2;
            sheet.Cells.SetRowHeight(i, height);
        }

        // Save the workbook to a file
        workbook.Save("RowHeights.xlsx");
    }
}
