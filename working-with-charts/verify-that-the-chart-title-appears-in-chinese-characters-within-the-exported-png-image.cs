// Title: Create a column chart with a Chinese title and export it to a PNG image using Aspose.Cells for .NET
// Common Searches: how to set a chart title in Chinese with Aspose.Cells C# | export Aspose.Cells chart to PNG with Unicode title | Aspose.Cells C# column chart Chinese characters title verification | render chart to high‑resolution PNG and check file size using Aspose.Cells | validate that chart title text is preserved after exporting to PNG in .NET
// Tags: unicode chart title Aspose.Cells | png chart rendering Aspose.Cells | column chart generation Aspose.Cells C# | image options dpi Aspose.Cells | verify chart title property C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace VerifyChartTitleChinese
{
    // The example creates a workbook, adds sample data, builds a column chart, assigns the Chinese string "销售额统计" as the chart title, confirms the title property, renders the chart to a 300 dpi PNG file, and then checks that the PNG file exists, is non‑empty, and can be read.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for the output PNG image
            const string outputImagePath = "ChartWithChineseTitle.png";

            // Chinese title text to be used
            const string chineseTitle = "销售额统计";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 2. Populate some sample data for the chart
                // ------------------------------------------------------------
                sheet.Cells["A1"].PutValue("月份");
                sheet.Cells["A2"].PutValue("一月");
                sheet.Cells["A3"].PutValue("二月");
                sheet.Cells["A4"].PutValue("三月");

                sheet.Cells["B1"].PutValue("销售额");
                sheet.Cells["B2"].PutValue(1500);
                sheet.Cells["B3"].PutValue(2300);
                sheet.Cells["B4"].PutValue(1800);

                // ------------------------------------------------------------
                // 3. Add a column chart and bind the data range
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // 4. Set the chart title to Chinese characters
                // ------------------------------------------------------------
                chart.Title.Text = chineseTitle;
                chart.Title.IsVisible = true;

                // Verify that the title property indeed contains the Chinese text
                if (chart.Title.Text != chineseTitle)
                {
                    Console.WriteLine("ERROR: Chart title was not set correctly.");
                    return;
                }
                Console.WriteLine($"Chart title set to Chinese text: \"{chart.Title.Text}\"");

                // ------------------------------------------------------------
                // 5. Export the chart to a PNG image
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true
                };

                // Save the chart image to file
                chart.ToImage(outputImagePath, imgOptions);
                Console.WriteLine($"Chart exported to PNG file: {outputImagePath}");

                // ------------------------------------------------------------
                // 6. Verify that the PNG file exists and has content
                // ------------------------------------------------------------
                if (!File.Exists(outputImagePath))
                {
                    Console.WriteLine("ERROR: PNG file was not created.");
                    return;
                }

                FileInfo fi = new FileInfo(outputImagePath);
                if (fi.Length == 0)
                {
                    Console.WriteLine("ERROR: PNG file is empty.");
                    return;
                }
                Console.WriteLine($"PNG file size: {fi.Length} bytes");

                // ------------------------------------------------------------
                // 7. (Optional) Simple visual verification by reading file bytes
                // ------------------------------------------------------------
                // Loading the image via System.Drawing is omitted to avoid extra dependencies.
                // Instead, we confirm that the file can be read.
                try
                {
                    byte[] imageBytes = File.ReadAllBytes(outputImagePath);
                    Console.WriteLine($"Read PNG image bytes successfully, length: {imageBytes.Length}");
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"WARNING: Unable to read PNG file bytes. {readEx.Message}");
                }

                // ------------------------------------------------------------
                // 8. Final confirmation
                // ------------------------------------------------------------
                Console.WriteLine("Verification completed: the chart title is set to Chinese characters and the PNG image was generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
