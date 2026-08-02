// Title: Aspose.Cells .NET – Apply a Custom RGB Palette to Chart Series
// Description: Creates a workbook, adds sample data, builds a column chart, defines three RGB colors, updates the workbook palette, assigns each color to a series via the Area.ForegroundColor property, forces solid fill, and saves the file as CustomPaletteChart.xlsx.
// Keywords: Aspose.Cells chart custom colors | C# set series RGB | change workbook palette Aspose | solid fill chart series | column chart color palette .NET
// Common Searches: set custom RGB colors for Aspose.Cells chart series | update workbook palette in Aspose.Cells C# | apply solid fill to chart series Aspose.Cells | change series area foreground color Aspose.Cells | custom color palette for Excel charts using Aspose
// Developer Intent: Generate a column chart and apply distinct RGB colors to each series by modifying the workbook palette.
// Use Cases: Brand‑consistent sales charts where each product line uses its corporate RGB shade. | Financial dashboards that enforce a corporate color scheme across multiple Excel charts. | Automated report generation that matches presentation templates with predefined palette colors.
// AI Prompts: Show how to extend the example to five series in a line chart with new RGB values. | Explain how to restore the default workbook palette after applying a custom one. | Provide code to use gradient fills for chart series while keeping a custom color palette.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsCustomChartPalette
{
    // Creates a workbook, adds sample data, builds a column chart, defines three RGB colors, updates the workbook palette, assigns each color to a series via the Area.ForegroundColor property, forces solid fill, and saves the file as CustomPaletteChart.xlsx.
    public class ApplyCustomPalette
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for three series
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                // Series 1
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Series 3
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(12);
                sheet.Cells["D3"].PutValue(22);
                sheet.Cells["D4"].PutValue(32);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the series (B1:D4) and categories (A2:A4)
                chart.NSeries.Add("B1:D4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Define custom colors (RGB)
                Color[] customColors = new Color[]
                {
                    Color.FromArgb(79, 129, 189),   // Series1
                    Color.FromArgb(192, 80, 77),    // Series2
                    Color.FromArgb(155, 187, 89)    // Series3
                };

                // Update the workbook palette so the custom colors are recognized
                for (int i = 0; i < customColors.Length; i++)
                {
                    workbook.ChangePalette(customColors[i], i);
                }

                // Apply the custom colors to each series via the Area.ForegroundColor property
                chart.NSeries[0].Area.ForegroundColor = customColors[0];
                chart.NSeries[1].Area.ForegroundColor = customColors[1];
                chart.NSeries[2].Area.ForegroundColor = customColors[2];

                // Ensure each series uses solid fill
                foreach (Series s in chart.NSeries)
                {
                    s.Area.Formatting = FormattingType.Custom;
                }

                // Save the workbook
                workbook.Save("CustomPaletteChart.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as CustomPaletteChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCustomPalette.Run();
        }
    }
}
