// Title: Set a solid light gray fill with 80 % opacity on a chart’s plot area using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart plot area fill color with transparency | How to apply 80% opacity solid fill to an Excel chart using Aspose.Cells | C# example for chart fill format solid color and transparency in Aspose.Cells | Set solid fill for column chart plot area in Aspose.Cells .NET | Aspose.Cells chart shading light gray with 20% transparency code
// Tags: chart plot area solid fill Aspose.Cells | fillformat transparency Aspose.Cells C# | set chart fill color Aspose.Cells .NET | excel chart shading Aspose.Cells | solid fill lightgray Aspose.Cells chart

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartFillExample
{
    // The sample creates a workbook, adds a column chart with sample data, and uses the FillFormat API to set the chart's plot area to a solid LightGray fill with 20% transparency (80% opacity), then saves the file as ChartWithSolidFill.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(45);
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(25);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ----- Set chart fill format to a solid color with 80% opacity -----
                // Use the PlotArea's Area to access the FillFormat
                FillFormat fill = chart.PlotArea.Area.FillFormat;

                // Solid fill
                fill.FillType = FillType.Solid;

                // Subtle shading color
                fill.SolidFill.Color = Color.LightGray;

                // 80% opacity => 20% transparency (0 = opaque, 1 = fully transparent)
                fill.SolidFill.Transparency = 0.20;

                // Output settings for verification
                Console.WriteLine("Chart fill type: " + fill.FillType);
                Console.WriteLine("Chart fill color: " + fill.SolidFill.Color);
                Console.WriteLine("Chart fill transparency (0‑1): " + fill.SolidFill.Transparency);

                // Save the workbook
                string outputPath = "ChartWithSolidFill.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
