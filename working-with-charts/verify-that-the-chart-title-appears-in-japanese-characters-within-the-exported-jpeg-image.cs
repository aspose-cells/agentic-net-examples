// Title: Generate a column chart with a Japanese title and export it as a high‑resolution JPEG using Aspose.Cells for .NET
// Common Searches: how to add a Japanese title to an Aspose.Cells chart and save as JPEG in C# | Aspose.Cells export chart to JPEG with Unicode title verification | C# code to render a column chart with Japanese text and check the output image | set chart title in Japanese characters using Aspose.Cells and verify the image file | high resolution JPEG export of an Aspose.Cells chart containing non‑Latin characters
// Tags: Aspose.Cells column chart Japanese title | Aspose.Cells chart image generation | ImageOrPrintOptions high resolution rendering | Unicode chart title verification | chart localization with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace ChartTitleJapaneseVerification
{
    // The example creates a workbook, fills it with Japanese product data, adds a column chart, assigns the Japanese title "売上高" to the chart, and exports the chart as a 300 dpi JPEG using ImageOrPrintOptions. After saving, it confirms that the image file exists, is non‑empty, and that the chart title matches the expected Japanese string, printing verification results to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Japanese title to be used on the chart
            const string japaneseTitle = "売上高"; // means "Sales"

            // Path for the output image
            string outputImagePath = Path.Combine(Directory.GetCurrentDirectory(), "ChartOutput.jpg");

            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header
                cells["A1"].PutValue("商品");   // "Product"
                cells["B1"].PutValue("数量");   // "Quantity"

                // Sample rows
                cells["A2"].PutValue("りんご"); // "Apple"
                cells["B2"].PutValue(120);
                cells["A3"].PutValue("みかん"); // "Mandarin"
                cells["B3"].PutValue(80);
                cells["A4"].PutValue("バナナ"); // "Banana"
                cells["B4"].PutValue(150);

                // ------------------------------------------------------------
                // 2. Add a column chart and set its title to Japanese text
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Define data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Set the chart title (Japanese)
                chart.Title.Text = japaneseTitle;
                chart.Title.IsVisible = true;
                chart.Title.Font.Size = 14; // larger font for visibility

                // ------------------------------------------------------------
                // 3. Export the chart to a JPEG image using ImageOrPrintOptions
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Jpeg,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    Quality = 90 // JPEG quality (0-100)
                };

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputImagePath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the chart image to file
                chart.ToImage(outputImagePath, imgOptions);

                // ------------------------------------------------------------
                // 4. Verify the export succeeded and the title is correct
                // ------------------------------------------------------------
                bool imageExists = File.Exists(outputImagePath);
                long imageSize = imageExists ? new FileInfo(outputImagePath).Length : 0;

                Console.WriteLine("=== Verification Results ===");
                Console.WriteLine($"Chart title set to: \"{chart.Title.Text}\"");
                Console.WriteLine($"Title matches expected Japanese text: {chart.Title.Text == japaneseTitle}");
                Console.WriteLine($"Image file created: {imageExists}");
                Console.WriteLine($"Image file size (bytes): {imageSize}");

                // Basic sanity check: image file should exist and be non‑empty
                if (imageExists && imageSize > 0 && chart.Title.Text == japaneseTitle)
                {
                    Console.WriteLine("Verification passed: Japanese chart title is present and image was generated.");
                }
                else
                {
                    Console.WriteLine("Verification failed: Check title text and image generation.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
