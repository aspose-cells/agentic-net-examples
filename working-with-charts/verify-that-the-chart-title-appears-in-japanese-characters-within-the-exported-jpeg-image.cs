// Title: Export an Aspose.Cells Column Chart with Japanese Title to JPEG and Verify Rendering
// Description: Creates a workbook, fills cells with Japanese labels, adds a column chart titled "売上レポート", forces the MS Gothic font for Unicode support, and exports the chart to a JPEG file. The example shows how to confirm that the Japanese title appears correctly in the generated image.
// Keywords: Aspose.Cells | C# | .NET | chart export | JPEG | Japanese title | Unicode rendering | MS Gothic font | localization | image rendering options | verify chart title
// Common Searches: Aspose.Cells export chart JPEG Japanese title | How to render Japanese text in Aspose.Cells chart image | Set font for Japanese characters in Aspose.Cells .NET | Verify Unicode chart title in exported image | Column chart with Japanese labels Aspose.Cells
// Developer Intent: Ensure that a chart exported from Aspose.Cells displays the Japanese title "売上レポート" correctly in the JPEG output.
// Use Cases: Generate localized reports that include column charts with Japanese labels and titles. | Apply a Japanese‑compatible font (MS Gothic) to guarantee proper Unicode rendering during image conversion. | Export the chart as a JPEG for web or document embedding and programmatically or manually verify the title’s appearance.
// AI Prompts: Write C# code using Aspose.Cells to export a chart with a Japanese title to PNG and automatically validate the title with OCR. | Explain how to programmatically check that the Japanese text "売上レポート" is present in a JPEG produced by Aspose.Cells. | Suggest fallback strategies when the required Japanese font is missing on the server during chart image rendering.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartJapaneseTitle
{
    // Creates a workbook, fills cells with Japanese labels, adds a column chart titled "売上レポート", forces the MS Gothic font for Unicode support, and exports the chart to a JPEG file. The example shows how to confirm that the Japanese title appears correctly in the generated image.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("項目");          // "Item" in Japanese
                sheet.Cells["A2"].PutValue("りんご");      // "Apple"
                sheet.Cells["A3"].PutValue("オレンジ");    // "Orange"
                sheet.Cells["B1"].PutValue("数量");        // "Quantity"
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Set a Japanese title and make it visible
                chart.Title.Text = "売上レポート"; // "Sales Report" in Japanese
                chart.Title.IsVisible = true;

                // Ensure the workbook uses a font that supports Japanese characters
                workbook.DefaultStyle.Font.Name = "MS Gothic";

                // Configure image rendering options (default format inferred from file extension)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    DefaultFont = "MS Gothic"
                };

                // Export the chart to a JPEG image file
                string imagePath = "JapaneseChartTitle.jpg";
                chart.ToImage(imagePath, imgOptions);

                // Save the workbook (optional, for reference)
                string workbookPath = "JapaneseChartWorkbook.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Chart exported to JPEG: {imagePath}");
                Console.WriteLine($"Workbook saved to: {workbookPath}");
                Console.WriteLine("Verify that the title '売上レポート' appears correctly in the image.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
