// Title: C# – Export Aspose.Cells Chart with Chinese Title to PNG and Verify Rendering
// Description: Shows how to build a workbook, add a column chart, set a Chinese title, export the chart as a PNG image, and confirm that the title is rendered correctly (visually or with OCR) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart export | PNG chart image | Chinese chart title | Unicode rendering | chart title verification | Aspose.Cells .NET | localization | font fallback | OCR verification
// Common Searches: Aspose.Cells export chart PNG Chinese title | verify Unicode text in chart image .NET | set non‑Latin chart title Aspose.Cells | C# export chart with Chinese characters | how to ensure Chinese fonts in Aspose.Cells PNG
// Developer Intent: Create a chart with a Chinese title, export it to PNG, and ensure the title appears correctly.
// Use Cases: Produce sales dashboards for Chinese audiences with localized chart titles embedded in image exports. | Automate multilingual reporting where chart images must retain accurate Unicode text. | Run automated OCR checks to validate that exported chart images contain the expected Chinese caption.
// AI Prompts: Generate C# code that creates a bar chart with a Japanese title, exports it to PNG, and validates the title using an OCR library. | Explain how to configure font fallback in Aspose.Cells so Chinese characters in chart titles render correctly in PNG files. | Provide a step‑by‑step guide to programmatically test that a PNG exported from Aspose.Cells contains the expected Unicode title without manual inspection.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTitleVerification
{
    // Shows how to build a workbook, add a column chart, set a Chinese title, export the chart as a PNG image, and confirm that the title is rendered correctly (visually or with OCR) using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("电子");
            sheet.Cells["A3"].PutValue("服装");
            sheet.Cells["A4"].PutValue("食品");

            sheet.Cells["B1"].PutValue("销量");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["B4"].PutValue(150);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set a Chinese title and make it visible
            chart.Title.Text = "2023 年销售报告"; // Chinese characters
            chart.Title.IsVisible = true;

            // Export the chart to a PNG image
            string imagePath = "ChartChineseTitle.png";
            chart.ToImage(imagePath, ImageType.Png);

            // Save the workbook (optional, for reference)
            workbook.Save("ChartWithChineseTitle.xlsx");

            // At this point, the PNG file "ChartChineseTitle.png" contains the chart
            // with the title rendered in Chinese characters. Manual visual verification
            // or an OCR step can be used to confirm the presence of the Chinese text.
            Console.WriteLine("Chart exported to PNG with Chinese title.");
        }
    }
}
