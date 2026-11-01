// Title: Generate an Excel workbook with a summary sheet of Aspose.Cells theme color names and their hexadecimal values using C#
// Common Searches: C# Aspose.Cells write a worksheet that shows theme color names and hex values | How to export default theme colors to an Excel file with Aspose.Cells .NET | Aspose.Cells create a color summary sheet with HTML hex codes in C# | List built‑in theme colors and their hex representations using Aspose.Cells | Generate ThemeColorsSummary.xlsx using fallback theme colors in Aspose.Cells
// Tags: Aspose.Cells create theme color summary sheet | C# export theme color hex values to Excel | fallback theme colors enumeration Aspose.Cells | System.Drawing.Color to HTML hex conversion Aspose.Cells | generate ThemeColorsSummary.xlsx programmatically

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeColorsSummary
{
    // The example creates a new Workbook, adds a worksheet named "ThemeColorsSummary", writes column headers for theme color name and hex value, iterates a predefined fallback list of six theme colors (Accent1‑Accent6), converts each System.Drawing.Color to its HTML hex string, records the name and hex code in successive rows, ensures the output directory exists, and saves the workbook as ThemeColorsSummary.xlsx while handling potential exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (default theme is applied)
                Workbook workbook = new Workbook();

                // Add a worksheet for the summary
                int sheetIndex = workbook.Worksheets.Add();
                Worksheet summarySheet = workbook.Worksheets[sheetIndex];
                summarySheet.Name = "ThemeColorsSummary";

                // Write header titles
                summarySheet.Cells[0, 0].PutValue("Theme Color Name");
                summarySheet.Cells[0, 1].PutValue("Hex Value");

                int currentRow = 1;

                // -----------------------------------------------------------------
                // NOTE:
                // The ThemeColorScheme API is not available in older Aspose.Cells
                // versions. To keep the example functional across versions, we
                // provide a fallback list of common theme colors.
                // -----------------------------------------------------------------
                var fallbackColors = new (string Name, Color Color)[]
                {
                    ("Accent1", Color.FromArgb(0, 112, 192)),
                    ("Accent2", Color.FromArgb(255, 192, 0)),
                    ("Accent3", Color.FromArgb(112, 173, 71)),
                    ("Accent4", Color.FromArgb(255, 0, 0)),
                    ("Accent5", Color.FromArgb(0, 176, 80)),
                    ("Accent6", Color.FromArgb(255, 0, 255))
                };

                foreach (var (name, color) in fallbackColors)
                {
                    summarySheet.Cells[currentRow, 0].PutValue(name);
                    string hex = ColorTranslator.ToHtml(color);
                    summarySheet.Cells[currentRow, 1].PutValue(hex);
                    currentRow++;
                }

                // Define output file path
                string outputPath = "ThemeColorsSummary.xlsx";

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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
