// Title: How to set incremental row heights for a range of rows using Aspose.Cells in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to loop through rows 0‑9 and set each row height increasing by 2 points starting from 15 points. | Show how to call Worksheet.Cells.SetRowHeight inside a for‑loop to adjust several consecutive rows and then save the workbook as an .xlsx file. | Provide an example that calculates row height based on the row index and applies it with Aspose.Cells for .NET.
// Common Searches: asp.net set different row heights in Excel using Aspose.Cells loop | c# Aspose.Cells increase row height by 2 points for each row | how to programmatically adjust multiple row heights in an Excel file with Aspose.Cells .NET
// Tags: Aspose.Cells SetRowHeight loop | C# incremental Excel row height | Worksheet.Cells.SetRowHeight multiple rows | Aspose.Cells dynamic row height .xlsx | programmatic row height adjustment .NET

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, looping through ten rows, and using Worksheet.Cells.SetRowHeight to increase each row's height by 2 points starting from 15 points, then saving the file as RowHeights.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Starting row index (0‑based), number of rows to modify, and base height
        int startRow = 0;
        int rowCount = 10;
        double baseHeight = 15.0; // points

        // Loop through the rows and set incremental heights
        for (int i = 0; i < rowCount; i++)
        {
            // Increase height by 2 points for each subsequent row
            double height = baseHeight + i * 2.0;

            // Apply the height to the current row
            sheet.Cells.SetRowHeight(startRow + i, height);
        }

        // Save the workbook to a file
        workbook.Save("RowHeights.xlsx");
    }
}
