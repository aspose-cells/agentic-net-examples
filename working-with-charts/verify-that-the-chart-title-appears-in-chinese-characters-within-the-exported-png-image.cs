// Title: C# – Export Aspose.Cells Chart with Chinese Title to PNG and Verify Rendering
// Description: Shows how to build a workbook, add a column chart, assign a Chinese‑language title, and export the chart as a PNG with Aspose.Cells for .NET, so you can confirm that the Unicode title appears correctly in the image.
// Keywords: Aspose.Cells | C# chart export | PNG chart image | Chinese title | Unicode chart text | non‑Latin characters | .NET localization | font fallback Aspose | chart title verification
// Common Searches: Aspose.Cells export chart with Chinese title | C# create chart and save as PNG Aspose | verify Unicode text in exported chart image | how to set non‑Latin chart title in Aspose.Cells | chart title rendering issue Chinese characters
// Developer Intent: The developer needs to ensure that a chart title containing Chinese characters is rendered correctly when the chart is exported to a PNG file using Aspose.Cells for .NET.
// Use Cases: Generate localized sales dashboards where chart headings must display Chinese text in image assets. | Automate creation of multilingual reports that embed chart PNGs in PDFs or web pages. | Perform visual or OCR‑based validation of chart rendering for quality‑assurance pipelines.
// AI Prompts: Write C# code with Aspose.Cells to export a chart that has a Japanese (or any Unicode) title to PNG and confirm the text appears. | Explain how to set up font fallback in Aspose.Cells so Chinese characters render properly in chart titles for image exports. | Suggest a programmatic way to verify the exported PNG contains the expected Chinese title using an OCR library.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartChineseTitleVerification
{
    // Shows how to build a workbook, add a column chart, assign a Chinese‑language title, and export the chart as a PNG with Aspose.Cells for .NET, so you can confirm that the Unicode title appears correctly in the image.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("水果");
            sheet.Cells["A3"].PutValue("蔬菜");
            sheet.Cells["B1"].PutValue("数量");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Set the chart title to Chinese characters and make it visible
            chart.Title.Text = "销售报告";          // "Sales Report" in Chinese
            chart.Title.IsVisible = true;

            // Export the chart to a PNG image
            string imagePath = "ChartChineseTitle.png";
            chart.ToImage(imagePath, ImageType.Png);

            // Save the workbook (optional, for reference)
            workbook.Save("ChartWithChineseTitle.xlsx");

            // Simple verification: read back the title text from the chart object
            // (The visual verification of Chinese characters in the PNG should be done manually or via OCR)
            Console.WriteLine("Chart title set to: " + chart.Title.Text);
            Console.WriteLine("Chart image saved to: " + imagePath);
        }
    }
}
