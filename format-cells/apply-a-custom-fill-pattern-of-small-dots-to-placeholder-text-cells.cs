// Title: How to apply a solid fill pattern to Excel cells containing {{placeholder}} text using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set solid background for cells that contain double curly braces | Apply custom fill pattern to placeholder strings in an existing workbook with Aspose.Cells | Iterate used range and change cell style based on placeholder text in C# | How to change foreground and background colors of specific Excel cells using Aspose.Cells | C# code to style cells with {{}} markers in an XLSX file using Aspose.Cells
// Tags: Aspose.Cells apply solid fill pattern | C# conditional cell styling placeholder text | modify cell style used range Aspose.Cells | set foreground background colors Excel .NET | placeholder cell formatting Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook with Aspose.Cells, scans every used cell, detects strings that contain "{{" and "}}", and applies a solid fill pattern with black foreground and white background colors before saving the updated file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load existing workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];

                    // Identify placeholder text wrapped in {{ }}
                    if (cell.Type == CellValueType.IsString &&
                        cell.StringValue.Contains("{{") && cell.StringValue.Contains("}}"))
                    {
                        // Get current style
                        var style = cell.GetStyle();

                        // Apply a fill pattern (using a valid BackgroundType)
                        style.Pattern = BackgroundType.Solid; // fallback to a common pattern

                        style.ForegroundColor = Color.Black;   // Pattern color
                        style.BackgroundColor = Color.White;   // Background color

                        // Set the modified style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
