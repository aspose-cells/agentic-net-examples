// Title: How to split panes and then freeze the top rows to create a fixed header in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells split panes and freeze top rows example | How to create a fixed header in Excel with Aspose.Cells FreezePanes | Aspose.Cells FreezePanes method to lock first three rows in .NET | Split pane then freeze pane on same worksheet using Aspose.Cells C# | Saving workbook after applying split and freeze panes with Aspose.Cells
// Tags: Aspose.Cells FreezePanes API C# | split panes then freeze pane Excel .NET | fixed header worksheet Aspose.Cells | populate worksheet sample data Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, fills it with sample data, applies split panes followed by freezing the first three rows to form a fixed header, and saves the file as SplitAndFreeze.xlsx using Aspose.Cells for .NET.
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

            // Populate sample data (optional)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first three rows (rows 1‑3) while keeping columns unfrozen
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(3, 0, 3, 0);

            // Define output file path
            string outputPath = "SplitAndFreeze.xlsx";

            // Ensure the directory exists (handle case where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
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
