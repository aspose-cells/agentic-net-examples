// Title: Add a caption below a chart and export it as PNG using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells add text label under chart before exporting to PNG | How to embed a caption in an Excel chart image with Aspose.Cells | Export Excel chart to PNG with a custom caption using Aspose.Cells C# | Shapes.AddLabelInChart example for chart caption in C# | Add external caption below chart image using Aspose.Cells rendering
// Tags: Aspose.Cells add label in chart | export chart to PNG Aspose.Cells | chart caption rendering C# | Shapes.AddLabelInChart usage | ImageOrPrintOptions chart export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AddCaptionToChartImage
{
    // The sample loads or creates an Excel workbook, inserts a column chart, adds a label inside the chart area as a caption using Shapes.AddLabelInChart, renders the chart to a PNG with ImageOrPrintOptions, and saves the final image with the caption embedded.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input workbook and output image
            string excelPath = "input.xlsx";
            string chartImagePath = "chart_with_caption.png";

            try
            {
                // -------------------------------------------------------------
                // 1. Load existing workbook or create a sample one with a chart
                // -------------------------------------------------------------
                Workbook workbook;
                if (File.Exists(excelPath))
                {
                    workbook = new Workbook(excelPath);
                }
                else
                {
                    workbook = CreateSampleWorkbook();
                    workbook.Save(excelPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Sample workbook created at '{excelPath}'.");
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No chart found in the worksheet.");
                    return;
                }

                Chart chart = sheet.Charts[0]; // the chart we added in the sample workbook

                // -------------------------------------------------------------
                // 2. Add a caption label inside the chart (appears below the plot area)
                // -------------------------------------------------------------
                // Parameters are in 1/4000 of the chart area. Adjust as needed.
                const int top = 3500;    // position from top
                const int left = 0;      // position from left
                const int height = 500;  // height of the label
                const int width = 4000;  // width spanning the chart

                var label = chart.Shapes.AddLabelInChart(top, left, height, width);
                label.Text = "Sample Chart Caption";
                label.Font.Size = 12;
                label.Font.Color = System.Drawing.Color.Black;
                label.Font.IsBold = false;

                // -------------------------------------------------------------
                // 3. Export the chart (with caption) to a PNG image
                // -------------------------------------------------------------
                var imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 96,
                    VerticalResolution = 96,
                    Transparent = false
                };

                // Export chart to temporary image
                string tempChartPath = "temp_chart.png";
                chart.ToImage(tempChartPath, imgOptions);

                // -------------------------------------------------------------
                // 4. (Optional) Add external caption below the chart image
                // -------------------------------------------------------------
                // In this example we simply copy the temporary image to the final path.
                // The caption is already inside the chart image.
                AddExternalCaption(tempChartPath, chartImagePath, "External Caption Below Chart");

                // Clean up temporary file
                if (File.Exists(tempChartPath))
                {
                    File.Delete(tempChartPath);
                }

                Console.WriteLine($"Chart with caption exported to '{chartImagePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = ws.Charts[chartIndex];
            chart.SetChartDataRange("A1:B4", true);
            chart.Title.Text = "Sample Column Chart";

            return wb;
        }

        private static void AddExternalCaption(string sourceImagePath, string outputImagePath, string captionText)
        {
            if (!File.Exists(sourceImagePath))
                throw new FileNotFoundException("Source chart image not found.", sourceImagePath);

            // For simplicity, just copy the source image to the output path.
            // The caption is already embedded inside the chart image.
            File.Copy(sourceImagePath, outputImagePath, true);
        }
    }
}
