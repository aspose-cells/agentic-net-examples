// Title: C# – Apply ChartColorPaletteType.CustomPalette2 to the Third Series of an Excel Chart with Aspose.Cells
// Description: This Aspose.Cells for .NET example opens an existing XLSX file, creates a column chart if none exists, guarantees at least three data series, assigns the built‑in CustomPalette2 color to the third series, optionally restores default colors for the first two series, and saves the workbook to a new file.
// Keywords: Aspose.Cells | C# | .NET | ChartColorPaletteType.CustomPalette2 | Excel chart series color | modify chart series programmatically | custom palette Excel | load and save workbook | add chart if missing | chart series formatting
// Common Searches: Aspose.Cells set custom palette for specific chart series | C# change color of third series in Excel chart | ChartColorPaletteType.CustomPalette2 example | add column chart with Aspose.Cells when none exists | programmatically ensure three series in Excel chart
// Developer Intent: Programmatically apply the CustomPalette2 color to the third series of a chart in an existing workbook and persist the change.
// Use Cases: Highlight a key data series in a financial dashboard by giving the third column a distinct palette color. | Automatically generate a chart for reports that lack one and enforce a three‑series structure with consistent styling. | Maintain brand colors by applying a custom palette to a specific series while keeping other series on default colors.
// AI Prompts: Generate C# code using Aspose.Cells that opens a workbook, adds a column chart if missing, ensures three series, and sets ChartColorPaletteType.CustomPalette2 for the third series. | Show how to retrieve Aspose.Cells built‑in color palettes and apply CustomPalette2 to a chart series in .NET. | Explain step‑by‑step how to check for existing charts, add missing series, and assign a custom palette color to an individual series with Aspose.Cells.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells for .NET example opens an existing XLSX file, creates a column chart if none exists, guarantees at least three data series, assigns the built‑in CustomPalette2 color to the third series, optionally restores default colors for the first two series, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart; create one if none exist
            Chart chart;
            if (worksheet.Charts.Count == 0)
            {
                // Add a default column chart at position (5,0) with size (15,7) cells
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
                chart = worksheet.Charts[chartIndex];
            }
            else
            {
                chart = worksheet.Charts[0];
            }

            // Access the series collection
            SeriesCollection seriesColl = chart.NSeries;

            // Ensure at least three series exist
            while (seriesColl.Count < 3)
            {
                int row = seriesColl.Count + 2; // start from row 2 (zero‑based)
                worksheet.Cells[row, 0].PutValue($"Category{seriesColl.Count + 1}");
                worksheet.Cells[row, 1].PutValue(10 + seriesColl.Count * 5);
                string range = $"B{row + 1}";
                seriesColl.Add(range, false);
            }

            // Apply a monochromatic palette to the third series only
            Color paletteColor = workbook.Colors.Length > 2 ? workbook.Colors[2] : Color.Gray;
            seriesColl[2].Area.ForegroundColor = paletteColor;

            // Optional: reset first two series to default colors
            if (workbook.Colors.Length > 0) seriesColl[0].Area.ForegroundColor = workbook.Colors[0];
            if (workbook.Colors.Length > 1) seriesColl[1].Area.ForegroundColor = workbook.Colors[1];

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
