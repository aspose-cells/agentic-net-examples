// Title: Apply the Light2 theme background color to a specific data table range in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to apply Light2 theme color to a range of cells | set solid background using GetThemeColor for a data table in Excel with Aspose.Cells | apply theme color with fallback Accent1 Accent2 to a cell range in .NET | format A1:D10 with Light2 background using Aspose.Cells style flag | Aspose.Cells example for applying theme background to a table range
// Tags: apply Light2 theme background Aspose.Cells | solid background style flag C# | retrieve theme color Accent1 fallback Aspose.Cells | format Excel data table range Aspose.Cells | create and apply style to cell range .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, defines the A1:D10 range, retrieves the Light2 theme color (using Accent1 with a fallback to Accent2), creates a solid background style, applies it to all cells in the range with a StyleFlag that sets all attributes, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data table range (A1:D10)
            int startRow = 0;      // Row 1 (zero‑based)
            int startColumn = 0;   // Column A
            int endRow = 9;        // Row 10
            int endColumn = 3;     // Column D

            // Retrieve a theme color (using Accent1 as primary, fallback to Accent2)
            Color themeColor;
            try
            {
                themeColor = workbook.GetThemeColor(ThemeColorType.Accent1);
            }
            catch (Exception)
            {
                // Fallback to another known theme color
                themeColor = workbook.GetThemeColor(ThemeColorType.Accent2);
            }

            // Create a style that uses the retrieved background color
            Style style = workbook.CreateStyle();
            style.ForegroundColor = themeColor;
            style.Pattern = BackgroundType.Solid;

            // Apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Apply the style to the defined range
            int rowCount = endRow - startRow + 1;
            int columnCount = endColumn - startColumn + 1;
            try
            {
                sheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount).ApplyStyle(style, flag);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to apply style: {ex.Message}");
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
