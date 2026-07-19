// Title: Set chart series fill color with an RGB value using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, and safely applies a solid fill color (RGB 79,129,189) to the first series via the Area → FillFormat → SolidFill API, then saves the file as XLSX.
// Keywords: Aspose.Cells chart series fill color | set series RGB color Aspose.Cells | C# chart series solid fill | Color.FromArgb Aspose.Cells | column chart series color .NET | chart FillFormat Aspose.Cells
// Common Searches: how to change chart series color in Aspose.Cells C# | set RGB fill for a series in Aspose.Cells chart | Aspose.Cells solid fill format for chart series | apply corporate brand color to Aspose.Cells chart | null‑safe way to set series fill color Aspose.Cells
// Developer Intent: Apply a specific RGB solid‑fill color to a chart series.
// Use Cases: Match corporate branding by coloring chart series with exact RGB shades. | Highlight key data points by assigning custom colors to individual series. | Generate reports where series colors are driven by business rules or thresholds.
// AI Prompts: Write C# code that uses Aspose.Cells to assign different RGB solid‑fill colors to each series in a bar chart with null checks. | Show how to modify the FillFormat of a chart series in Aspose.Cells safely, including gradient fill examples. | Create a reusable method that sets a series' solid fill color from a hex string in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesFillColorDemo
{
    // Creates a workbook, adds sample data, inserts a column chart, and safely applies a solid fill color (RGB 79,129,189) to the first series via the Area → FillFormat → SolidFill API, then saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Safely set the fill color of the first series
                if (chart.NSeries.Count > 0 && chart.NSeries[0] != null)
                {
                    // Ensure the Area and FillFormat objects are instantiated
                    var seriesArea = chart.NSeries[0].Area;
                    if (seriesArea != null && seriesArea.FillFormat != null && seriesArea.FillFormat.SolidFill != null)
                    {
                        seriesArea.FillFormat.SolidFill.Color = Color.FromArgb(79, 129, 189);
                    }
                }

                // Save the workbook
                workbook.Save("SeriesFillColorDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
