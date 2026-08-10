// Title: Aspose.Cells for .NET – Export a Chart with a Japanese Title to JPEG
// Description: Creates a workbook, adds a column chart with Japanese category labels and a Japanese title, sets MS Gothic as the default font, exports the chart to a JPEG image, and saves the workbook. The JPEG can be inspected manually or with OCR to confirm that the title renders in Japanese characters.
// Keywords: Aspose.Cells | .NET | C# | chart export JPEG | Japanese chart title | Unicode rendering | MS Gothic font | ImageOrPrintOptions | chart verification | Japan locale | Excel chart to image
// Common Searches: export Aspose.Cells chart to JPEG with Japanese text | set default font for Japanese characters in Aspose chart image | verify Japanese title in exported chart image Aspose.Cells | Aspose.Cells Unicode support for chart titles | C# Aspose chart image Japanese characters
// Developer Intent: Ensure that a chart title containing Japanese characters is correctly rendered in the JPEG image produced by Aspose.Cells.
// Use Cases: Generate Excel charts for Japanese reports and export them as images without losing glyphs. | Automate visual or OCR‑based validation of multilingual chart titles in CI pipelines. | Create localized dashboards where chart titles must appear in native scripts.
// AI Prompts: Write C# code that adds a chart with a Japanese title to a workbook and exports it to PNG, handling font fallback automatically. | Explain how to programmatically confirm that a JPEG produced by Aspose.Cells contains the expected Japanese title using OCR libraries. | Suggest ImageOrPrintOptions settings that guarantee correct rendering of East Asian characters in chart images.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartJapaneseTitleVerification
{
    // Creates a workbook, adds a column chart with Japanese category labels and a Japanese title, sets MS Gothic as the default font, exports the chart to a JPEG image, and saves the workbook. The JPEG can be inspected manually or with OCR to confirm that the title renders in Japanese characters.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("項目");          // "Item" in Japanese
            sheet.Cells["A2"].PutValue("リンゴ");      // "Apple"
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
            chart.Title.Text = "日本語のタイトル"; // "Japanese Title"
            chart.Title.IsVisible = true;

            // Ensure the image rendering uses a font that supports Japanese characters
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,
                DefaultFont = "MS Gothic"   // Font that contains Japanese glyphs
            };

            // Export the chart to a JPEG image
            string imagePath = "ChartWithJapaneseTitle.jpg";
            chart.ToImage(imagePath, imgOptions);

            // Save the workbook (optional, for further inspection)
            workbook.Save("WorkbookWithJapaneseChart.xlsx");

            // At this point, the JPEG file "ChartWithJapaneseTitle.jpg" should display the title
            // in Japanese characters. Manual visual verification or OCR can be used to confirm.
            Console.WriteLine("Chart exported to JPEG with Japanese title. Verify the image at: " + imagePath);
        }
    }
}
