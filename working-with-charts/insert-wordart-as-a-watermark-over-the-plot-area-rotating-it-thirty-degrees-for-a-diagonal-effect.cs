using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class WordArtWatermarkInChart
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart covering a range of rows/columns
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Insert WordArt into the chart's plot area
                Shape wordArt = chart.Shapes.AddTextEffectInChart(
                    MsoPresetTextEffect.TextEffect2, // preset effect
                    "CONFIDENTIAL",                  // watermark text
                    "Arial Black",                   // font name
                    36,                              // font size
                    true,                            // bold
                    false,                           // italic
                    2000,                            // left offset (1/4000 of chart)
                    2000,                            // top offset
                    6000,                            // width
                    2000);                           // height

                // Rotate the WordArt for watermark effect
                wordArt.RotationAngle = -45;

                // Ensure output directory exists
                string outputPath = "WordArtWatermarkInChart.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
                throw;
            }
        }
    }
}