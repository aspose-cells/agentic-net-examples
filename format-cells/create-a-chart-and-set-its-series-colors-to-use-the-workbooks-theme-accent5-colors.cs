// Title: Create a column chart and apply the workbook’s Accent5 theme color to all series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart series fill color to workbook theme Accent5 | C# Aspose.Cells apply workbook theme color to column chart series | How to retrieve Accent5 theme color from workbook in Aspose.Cells | Set series border and marker colors from workbook theme using Aspose.Cells .NET | Create column chart with theme-based series styling in Aspose.Cells C#
// Tags: chart series theme color Aspose.Cells | series fill color using workbook theme | workbook theme color extraction C# | column chart series theme styling | theme-based series coloring Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // The example creates a new workbook, fills it with sample data, adds a column chart, retrieves the workbook's Accent5 theme color, and applies that color to each series' fill, border, and marker settings before saving the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart to the worksheet
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
                Chart chart = sheet.Charts[chartIdx];
                chart.Title.Text = "Sample Chart";

                // Define the series and category data
                chart.NSeries.Add("B2:B4", true); // Series1 values
                chart.NSeries.Add("C2:C4", true); // Series2 values
                chart.NSeries.CategoryData = "A2:A4";

                // Retrieve the workbook's Accent5 theme color
                Color accent5Color = workbook.GetThemeColor(ThemeColorType.Accent5);

                // Apply the Accent5 color to each series (fill, border, markers)
                foreach (Series series in chart.NSeries)
                {
                    // Fill color for the series area (bars/columns)
                    series.Area.ForegroundColor = accent5Color;

                    // Outline/border color (relevant for line charts, safe for column charts)
                    series.Border.Color = accent5Color;

                    // Marker colors (relevant for scatter/line charts)
                    series.Marker.ForegroundColor = accent5Color;
                    series.Marker.BackgroundColor = accent5Color;
                }

                // Save the workbook with the chart
                string outputPath = "ChartWithAccent5.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
