// Title: Change the workbook Accent2 theme color and confirm all chart legends use the new shade with Aspose.Cells for .NET
// Common Searches: how to update the Accent2 theme color and verify chart legend colors with Aspose.Cells in C# | Aspose.Cells C# check if chart legends inherit the modified theme accent color | C# code to iterate through all worksheets and charts to validate legend font color after SetThemeColor | validate workbook theme change reflects in chart legends using Aspose.Cells | example of using Workbook.SetThemeColor to affect chart legend appearance in .NET
// Tags: Workbook.SetThemeColor Accent2 Aspose.Cells | chart legend font color inheritance Aspose.Cells | enumerate worksheets and charts C# | theme color validation for Excel charts Aspose.Cells | update Excel theme accent color .NET | legend color verification after theme change

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The sample loads an existing Excel workbook, changes the Accent2 theme color via Workbook.SetThemeColor, then loops through every worksheet and each chart to compare the legend's font color with the new theme shade, logging mismatches or confirmations, and finally saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the new shade for a theme color (using Accent2 as an example)
            Color newAccent2Color = Color.FromArgb(255, 0, 120, 215); // example shade

            // Update the theme's Accent2 color
            workbook.SetThemeColor(ThemeColorType.Accent2, newAccent2Color);

            // Validate that each chart legend reflects the updated theme color
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Retrieve the legend's font color (legends inherit theme colors)
                    Color legendFontColor = chart.Legend.Font.Color;

                    // Compare with the newly set theme color
                    if (legendFontColor.ToArgb() != newAccent2Color.ToArgb())
                    {
                        Console.WriteLine($"[Mismatch] Worksheet: '{sheet.Name}', Chart: '{chart.Name}' - Legend color {legendFontColor} does not match updated theme color {newAccent2Color}.");
                    }
                    else
                    {
                        Console.WriteLine($"[OK] Worksheet: '{sheet.Name}', Chart: '{chart.Name}' - Legend correctly reflects the new theme color.");
                    }
                }
            }

            // Save the workbook (optional, to persist changes)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
