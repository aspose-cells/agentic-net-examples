// Title: Freeze first three columns and top four rows (split panes) in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to freeze columns A‑C and rows 1‑4 in a worksheet. | Show how to set split panes at cell D5 in a .NET workbook with Aspose.Cells. | Create a minimal example that saves an Excel file with frozen left columns and top rows using Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze panes first three columns and first four rows | how to set split panes at D5 using Aspose.Cells .NET API | Worksheet FreezePanes parameters row index column index explanation | example of saving Excel file with frozen top rows and left columns in C#
// Tags: Worksheet.FreezePanes left columns split pane | Aspose.Cells freeze panes C# example | Excel split panes D5 Aspose.Cells | freeze top rows and left columns Aspose.Cells .NET | save workbook with frozen panes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The program creates a new workbook, accesses the first worksheet, freezes the first three columns (A‑C) and the first four rows (1‑4) by calling FreezePanes(4, 3, 4, 3), and saves the file as SplitPanes.xlsx.
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

                // Freeze the first 3 columns (A‑C) and the first 4 rows (1‑4)
                // Freeze at cell D5 (row index 4, column index 3) and specify the number of rows/columns to freeze
                sheet.FreezePanes(4, 3, 4, 3);

                // Define output file path
                string outputPath = "SplitPanes.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
