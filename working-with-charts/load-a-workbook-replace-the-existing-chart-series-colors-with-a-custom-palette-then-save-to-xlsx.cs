// Title: Replace Excel chart series colors with a custom RGB palette using Aspose.Cells for .NET and save the workbook as XLSX
// AI Prompts: Write C# code that loads an existing XLSX file (or creates a workbook with a column chart if the file is missing), defines a Color[] palette, and assigns each chart series' Area.ForegroundColor and Border.Color to the palette colors using Aspose.Cells. | Provide a C# snippet that walks every worksheet and each chart, applies colors from a predefined palette to the series, and writes the updated workbook to a new XLSX file with Aspose.Cells.
// Common Searches: Aspose.Cells change series fill color in Excel chart C# | apply custom RGB palette to chart series Aspose.Cells .NET | C# iterate over workbook charts and set series border color Aspose.Cells | replace default chart colors with specific colors in Excel using Aspose.Cells | save modified workbook after updating chart colors Aspose.Cells
// Tags: Aspose.Cells chart series color customization | C# set chart series fill color Aspose.Cells | custom RGB palette Excel chart Aspose.Cells | iterate workbook charts Aspose.Cells .NET | save workbook as XLSX Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an existing XLSX workbook (or creates a simple one with a column chart if the file is absent), defines a six‑color RGB palette, iterates through all worksheets and their charts, and applies each palette color to the corresponding series' fill (Area.ForegroundColor) and border (Border.Color). Finally, it saves the modified workbook as 'output.xlsx' using Aspose.Cells for .NET.
class ReplaceChartSeriesColors
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook; if input file does not exist, create a simple workbook with a chart
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "SampleData";

                // Sample data
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = ws.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
            }

            // Custom palette of colors
            Color[] customPalette = new Color[]
            {
                Color.FromArgb(0x4F, 0x81, 0xBD), // Blue
                Color.FromArgb(0xC0, 0x50, 0x4D), // Red
                Color.FromArgb(0x9A, 0xB9, 0x59), // Green
                Color.FromArgb(0xFF, 0xC1, 0x00), // Yellow
                Color.FromArgb(0x5B, 0x9C, 0xD5), // Light Blue
                Color.FromArgb(0xFF, 0x99, 0x33)  // Orange
            };

            // Iterate through worksheets and charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    for (int s = 0; s < chart.NSeries.Count; s++)
                    {
                        Series series = chart.NSeries[s];
                        Color paletteColor = customPalette[s % customPalette.Length];

                        // Apply color to series area
                        series.Area.ForegroundColor = paletteColor;

                        // Apply color to series border (line)
                        if (series.Border != null)
                        {
                            series.Border.Color = paletteColor;
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
