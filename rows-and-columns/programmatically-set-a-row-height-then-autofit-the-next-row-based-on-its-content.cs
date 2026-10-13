// Title: Set a custom row height and auto‑fit the following row with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new Workbook, set row 0 height to 30 points, put a long string into cell A2, auto‑fit row 1, and save the file as Result.xlsx using Aspose.Cells in C#. | Using Aspose.Cells for .NET, programmatically assign a specific height to one row and then automatically adjust the height of the next row based on its cell content.
// Common Searches: Aspose.Cells C# set specific row height and auto fit next row | how to auto‑fit a row after setting custom height for another row using Aspose.Cells .NET | example of setting row height then auto‑fitting another row in C# with Aspose.Cells | Aspose.Cells programmatic row height adjustment followed by auto‑fit of subsequent row
// Tags: set row height Aspose.Cells .NET | auto fit row based on cell content Aspose.Cells | custom row height then auto fit next row C# | Aspose.Cells row height manipulation example | auto fit row after inserting long text C#

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, setting the first row height to 30 points, adding long text to the second row, auto‑fitting that row, and saving the workbook as Result.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a custom height (in points) for the first row (index 0)
        sheet.Cells.SetRowHeight(0, 30);

        // Add content to the next row (index 1) that will require auto‑fitting
        sheet.Cells["A2"].PutValue("This is a long piece of text that should cause the row height to expand when auto‑fitted.");

        // Auto‑fit the second row based on its content
        sheet.AutoFitRow(1);

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}
