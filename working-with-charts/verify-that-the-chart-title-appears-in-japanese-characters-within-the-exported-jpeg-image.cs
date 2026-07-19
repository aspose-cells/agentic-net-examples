// Title: Validate Japanese Chart Title in JPEG Export with Aspose.Cells for .NET (C#)
// Description: A C# example that builds a workbook, inserts Japanese labels, creates a column chart, assigns a Japanese title, selects a glyph‑compatible font, and renders the chart to a JPEG file so you can confirm the title appears correctly.
// Keywords: Aspose.Cells C# chart export | JPEG chart rendering | Japanese text in chart title | MS Gothic font Aspose.Cells | multilingual chart image | localization chart export | GitHub Aspose.Cells example
// Common Searches: export Aspose.Cells chart to JPEG with Japanese characters | which font supports Japanese glyphs in Aspose.Cells images | how to render non‑Latin titles in Aspose.Cells chart screenshots | C# verify chart title language after image conversion | Aspose.Cells chart localization tutorial
// Developer Intent: Ensure that a chart saved as a JPEG displays its title using Japanese characters.
// Use Cases: Produce sales dashboards for Japanese markets and embed chart images in reports. | Automate multilingual chart generation for web portals that require raster graphics. | Integrate visual validation of localized chart titles into CI/CD pipelines.
// AI Prompts: Generate C# code with Aspose.Cells that exports a chart containing a Japanese title to PNG and extracts the title text using OCR for automated verification. | Explain how to configure fallback fonts for East Asian scripts in Aspose.Cells rendering options and why MS Gothic is a safe choice. | Suggest a CI step that programmatically checks a JPEG chart for the expected Japanese title without manual inspection.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartJapaneseTitleVerification
{
    // A C# example that builds a workbook, inserts Japanese labels, creates a column chart, assigns a Japanese title, selects a glyph‑compatible font, and renders the chart to a JPEG file so you can confirm the title appears correctly.
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

                // Set Japanese title and make it visible
                chart.Title.Text = "売上レポート"; // "Sales Report" in Japanese
                chart.Title.IsVisible = true;

                // Configure image rendering options to use a font that supports Japanese characters
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; if JPEG is required, set ImageFormat when supported.
                    DefaultFont = "MS Gothic" // Font that contains Japanese glyphs
                };

                // Export the chart to an image file
                string imagePath = "ChartJapaneseTitle.jpg";
                chart.ToImage(imagePath, imgOptions);

                // Save the workbook (optional, for reference)
                string workbookPath = "ChartWithJapaneseTitle.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Chart image saved to '{Path.GetFullPath(imagePath)}'. Verify the Japanese title visually.");
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(workbookPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
