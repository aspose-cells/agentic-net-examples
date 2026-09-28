// Title: How to freeze columns A‑C and rows 1‑4 (split panes) in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, applies Worksheet.FreezePanes to lock the first three columns and first four rows, and saves the file. | Demonstrate how to verify or create the output folder before calling Workbook.Save with Aspose.Cells. | Explain the meaning of each parameter in Worksheet.FreezePanes for configuring split panes.
// Common Searches: Aspose.Cells C# example to freeze first three columns and first four rows | Worksheet.FreezePanes row column count parameters Aspose.Cells .NET | Create split panes in Excel using Aspose.Cells C# code | Save Excel workbook to a custom directory after freezing panes with Aspose.Cells
// Tags: freeze panes Worksheet.FreezePanes Aspose.Cells | split pane configuration Excel .NET | ensure output directory before Workbook.Save C# | freeze first three columns and four rows Aspose.Cells | create workbook and apply split panes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Shows how to generate a new workbook, freeze columns A‑C and rows 1‑4 with Worksheet.FreezePanes, ensure the target folder exists, and save the result as SplitPanes.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the pane so that the left frozen area includes the first three columns (A‑C)
            // and the top frozen area includes the first four rows (1‑4).
            // Parameters: row index, column index, total rows to freeze, total columns to freeze (zero‑based).
            sheet.FreezePanes(4, 3, 4, 3);

            // Define output file path.
            string outputPath = "SplitPanes.xlsx";

            // Ensure the directory for the output file exists.
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file.
            workbook.Save(fullOutputPath);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
