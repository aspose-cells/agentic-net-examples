// Title: Apply the Light2 Theme Background Color to a Data Table Range using Aspose.Cells for .NET (C#)
// Common Searches: how to set Light2 theme background color for a specific range in Excel with Aspose.Cells C# | Aspose.Cells C# apply solid fill from theme color to table A1:D10 | using Aspose.Cells .NET retrieve theme color and style a cell range | C# code to apply theme accent color as background to Excel range with Aspose.Cells | apply style flag all attributes to Excel range using Aspose.Cells
// Tags: apply theme background color Aspose.Cells C# | style range with Light2 theme Aspose.Cells | create solid fill style using workbook.GetThemeColor | range.ApplyStyle with StyleFlag All Aspose.Cells | set background for Excel data table Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, defines a rectangular range (e.g., A1:D10), retrieves the Light2 theme color via GetThemeColor, creates a solid‑fill style, applies it to the entire range with a StyleFlag that enables all style attributes, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Define the range of the data table (example: A1:D10)
                int startRow = 0;      // Row 1 (zero‑based)
                int startColumn = 0;   // Column A (zero‑based)
                int endRow = 9;        // Row 10
                int endColumn = 3;     // Column D

                // Retrieve a theme color (using Accent1 which is guaranteed to exist)
                Color themeColor = workbook.GetThemeColor(ThemeColorType.Accent1);

                // Create a style that uses the theme background color
                Style style = workbook.CreateStyle();
                style.ForegroundColor = themeColor;
                style.Pattern = BackgroundType.Solid;

                // Apply all style attributes
                StyleFlag flag = new StyleFlag { All = true };

                // Create the range and apply the style
                Aspose.Cells.Range range = sheet.Cells.CreateRange(
                    startRow,
                    startColumn,
                    endRow - startRow + 1,
                    endColumn - startColumn + 1);
                range.ApplyStyle(style, flag);

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
}
