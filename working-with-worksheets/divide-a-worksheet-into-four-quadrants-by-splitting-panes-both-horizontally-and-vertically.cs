// Title: How to split an Excel worksheet into four quadrants by freezing panes horizontally and vertically using Aspose.Cells for .NET
// AI Prompts: Write a C# program that creates a new workbook, freezes panes after row 15 and column 8 to produce four quadrants, and saves the file as QuadrantView.xlsx. | Modify the sample to split the sheet at row 25 and column 12, then save the workbook as SplitQuadrants.xlsx. | Detail how the four arguments of Worksheet.FreezePanes determine the position of the frozen area for quadrant splitting.
// Common Searches: Aspose.Cells C# how to freeze both rows and columns to get a quadrant layout | example code for dividing an Excel worksheet into four sections using FreezePanes in .NET | setting horizontal and vertical split points with Worksheet.FreezePanes Aspose.Cells | freeze first 10 rows and 5 columns in an Excel file using Aspose.Cells .NET | understanding FreezePanes parameters for creating pane quadrants in Excel
// Tags: Worksheet.FreezePanes quadrant layout | Aspose.Cells pane splitting API | C# generate four quadrant Excel view | programmatic Excel pane split .NET | freeze pane configuration Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, accesses the first worksheet, and calls Worksheet.FreezePanes to freeze the first 10 rows and 5 columns, thereby dividing the sheet into four quadrants. The workbook is saved as Quadrants.xlsx and the full output path is displayed.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze panes after row 10 and column 5 (zero‑based indices)
            // totalRows = 10, totalColumns = 5 to freeze the first 10 rows and 5 columns
            sheet.FreezePanes(10, 5, 10, 5);

            // Define output file name
            string outputPath = "Quadrants.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
