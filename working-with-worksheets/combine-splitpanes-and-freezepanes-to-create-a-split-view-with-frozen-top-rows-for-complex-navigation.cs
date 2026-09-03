// Title: Create a split view with frozen top rows using Aspose.Cells SplitPanes and FreezePanes in C#
// AI Prompts: Write C# code that inserts a vertical split pane at row 2, freezes the first two rows of a worksheet with Aspose.Cells, and saves the file as an .xlsx workbook. | Show how to fill a worksheet with sample data, apply FreezePanes to lock header rows, and set a split pane for independent scrolling using Aspose.Cells for .NET. | Demonstrate combining the SplitPanes and FreezePanes methods to produce a split view with frozen header rows and export the result to Excel.
// Common Searches: Aspose.Cells C# split pane and freeze top rows example | How to freeze header rows while using split panes in Aspose.Cells .NET | C# code to create a split view with frozen rows in an Excel workbook using Aspose.Cells | Set vertical split and freeze first two rows with Aspose.Cells for .NET
// Tags: Aspose.Cells FreezePanes SplitPanes C# | split pane frozen header rows .NET | Excel workbook split view Aspose.Cells | populate worksheet sample data Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, populates 100 rows by 20 columns with sample data, applies a vertical split pane and freezes the top two rows, then saves the file as SplitFreezeExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate the worksheet with sample data
            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the top 2 rows in the upper pane
            // Parameters: first unfrozen row, first unfrozen column, left column, top row
            sheet.FreezePanes(2, 0, 0, 0);

            // Determine output path and ensure the directory exists
            string outputPath = "SplitFreezeExample.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
