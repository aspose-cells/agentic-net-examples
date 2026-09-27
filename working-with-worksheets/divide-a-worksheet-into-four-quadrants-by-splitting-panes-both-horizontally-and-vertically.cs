// Title: Split an Excel worksheet into four quadrants by freezing panes horizontally and vertically with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to create a workbook, populate cells, and freeze panes at a given row and column to produce four quadrants. | Demonstrate how to call Worksheet.FreezePanes to split a sheet both horizontally and vertically and then set the active pane. | Adjust the split row and column values in an Aspose.Cells example to customize the size and position of the quadrants.
// Common Searches: Aspose.Cells C# freeze panes both rows and columns to create quadrant view | How to split an Excel worksheet into four panes using Aspose.Cells .NET | Set active pane after freezing rows and columns with Aspose.Cells | Example of Worksheet.FreezePanes for horizontal and vertical split in C# | Customize split positions for quadrants in Aspose.Cells workbook
// Tags: Worksheet.FreezePanes Aspose.Cells C# | split worksheet quadrants Aspose.Cells | set active pane Excel Aspose.Cells | freeze panes horizontally and vertically .NET | programmatic quadrant layout Excel C#

using System;
using System.IO;
using Aspose.Cells;

namespace SplitPaneExample
{
    // Creates a new workbook, fills a 20x10 range with sample data, freezes panes at row 10 and column 5 to divide the first worksheet into four quadrants, and saves the file as SplitPaneResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill some data to visualize the quadrants
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define split positions (zero‑based)
                int splitRow = 10;   // splits after row 10
                int splitColumn = 5; // splits after column 5

                // Split the worksheet and set the active pane.
                // Using FreezePanes without PaneType for compatibility with all versions.
                sheet.FreezePanes(splitRow, splitColumn, splitRow, splitColumn);

                // Define output file path
                string outputPath = "SplitPaneResult.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
