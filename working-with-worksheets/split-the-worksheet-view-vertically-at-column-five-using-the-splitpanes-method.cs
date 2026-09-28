// Title: How to split an Excel worksheet view vertically after column 5 using Aspose.Cells SplitPanes in C# (.NET)
// AI Prompts: Write C# code that creates a Workbook, accesses the first Worksheet, and applies a vertical split after the fifth column using Aspose.Cells SplitPanes (or FreezePanes) before saving the file as .xlsx. | Show how to verify or create the output directory, apply the vertical split, and wrap the operation in try‑catch error handling for a console application using Aspose.Cells. | Extend the sample to add a horizontal split at row 10 while keeping the existing vertical split at column 5, demonstrating multiple SplitPanes calls in C#.
// Common Searches: aspnet split pane after column 5 using Aspose.Cells | c# Aspose.Cells vertical split view column 5 example | how to freeze first five columns in Excel with Aspose.Cells .NET | Aspose.Cells SplitPanes method usage for vertical split | create split panes in workbook programmatically Aspose.Cells C#
// Tags: Aspose.Cells vertical split pane C# | SplitPanes method Excel .NET | freeze first columns worksheet Aspose.Cells | programmatic Excel view split Aspose.Cells | create output directory before saving Aspose.Cells | exception handling Aspose.Cells workbook save

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a new Workbook, retrieving the first Worksheet, and using Aspose.Cells' SplitPanes (or FreezePanes) method to split the worksheet view vertically after column five. It also shows how to ensure the target folder exists, save the workbook as an .xlsx file, and handle potential exceptions.
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

            // Freeze the first 5 columns (vertical split after column F)
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(0, 5, 0, 0);

            // Define output file path
            string outputPath = "SplitPaneExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
