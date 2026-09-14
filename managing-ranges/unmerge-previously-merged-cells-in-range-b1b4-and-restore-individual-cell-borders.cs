// Title: How to unmerge cells B1:B4 and add thin black borders to each cell using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel workbook, unmerges the range B1:B4, and applies thin black borders to every cell in that range. | Generate a .NET example that removes merging from column B rows 1‑4 and sets uniform thin black borders on each resulting cell using Aspose.Cells.
// Common Searches: Aspose.Cells C# unmerge cells in column B rows 1 to 4 | C# add thin black borders after unmerging a range with Aspose.Cells | How to restore individual cell borders after unmerge using Aspose.Cells for .NET | Unmerge B1:B4 and apply border style with Aspose.Cells in a C# project
// Tags: unmerge range B1:B4 Aspose.Cells | apply thin black borders Aspose.Cells C# | cell border styling after unmerge Aspose.Cells | Aspose.Cells range unmerge and style | C# Excel border formatting Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing Excel workbook, defines the B1:B4 range, unmerges any merged cells within that range, and then applies thin black borders to the top, bottom, left, and right of each individual cell before saving the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the range B1:B4 (rows 0-3, column 1)
            int startRow = 0;
            int endRow = 3;
            int column = 1; // Column B

            // Unmerge cells in the specified range (if any)
            try
            {
                int totalRows = endRow - startRow + 1;
                int totalColumns = 1;
                // Use fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range
                Aspose.Cells.Range range = sheet.Cells.CreateRange(startRow, column, totalRows, totalColumns);
                range.UnMerge();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to unmerge cells – {ex.Message}");
            }

            // Apply thin black borders to each individual cell in the range
            for (int row = startRow; row <= endRow; row++)
            {
                Cell cell = sheet.Cells[row, column];
                Style style = cell.GetStyle();

                // Set all four borders to thin black
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                style.Borders[BorderType.TopBorder].Color = Color.Black;
                style.Borders[BorderType.BottomBorder].Color = Color.Black;
                style.Borders[BorderType.LeftBorder].Color = Color.Black;
                style.Borders[BorderType.RightBorder].Color = Color.Black;

                cell.SetStyle(style);
            }

            // Ensure the output directory exists
            try
            {
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not create output directory – {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
