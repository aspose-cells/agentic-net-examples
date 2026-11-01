// Title: Combine SplitPanes and FreezePanes in Aspose.Cells for .NET to create a split view with frozen top rows
// AI Prompts: Generate C# code that uses Aspose.Cells to add a vertical split pane at a specific column, freeze the first two rows, and save the workbook as an .xlsx file. | Describe the steps for applying both SplitPanes and FreezePanes methods together in Aspose.Cells so that header rows stay visible while a split navigation pane is available.
// Common Searches: asp.net aspose.cells how to split pane and freeze top rows in Excel | c# create split view with frozen header rows using Aspose.Cells | example of FreezePanes combined with SplitPanes in Aspose.Cells .NET | set vertical split column and freeze first two rows Aspose.Cells C# tutorial
// Tags: Aspose.Cells SplitPanes API | Aspose.Cells FreezePanes usage | C# split pane with frozen rows Excel | Excel split view frozen header Aspose | Workbook navigation split pane .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, fills a 100‑row by 20‑column worksheet with data, applies FreezePanes(2,0,0,0) to lock the top two rows, adds a vertical split pane for navigation, and saves the file as SplitAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill the worksheet with sample data for navigation testing
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ---------- FreezePanes ----------
            // Freeze the top two rows so they remain visible while scrolling vertically.
            // The overload with four parameters is used for compatibility with newer API versions.
            sheet.FreezePanes(2, 0, 0, 0);

            // Define output file path
            string outputPath = "SplitAndFreeze.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
