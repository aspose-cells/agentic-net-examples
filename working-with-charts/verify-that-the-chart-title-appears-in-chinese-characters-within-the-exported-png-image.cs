// Title: Export a Column Chart with a Chinese Title to PNG using Aspose.Cells for .NET
// Description: Creates a workbook, fills cells with Chinese labels, adds a column chart, sets the title to "销售报告", and exports the chart as a PNG image. The PNG can be inspected manually or with OCR to confirm the Chinese characters are rendered correctly.
// Keywords: Aspose.Cells C# export chart PNG | chart title Chinese characters | Unicode chart title Aspose.Cells | column chart PNG Aspose.Cells | localize chart titles .NET
// Common Searches: Aspose.Cells export chart with Chinese title | C# chart PNG Unicode verification | how to show Chinese text in Aspose.Cells chart | export column chart to PNG with non‑Latin title | verify Chinese characters in chart image
// Developer Intent: Generate a PNG of a column chart whose title is displayed in Chinese and ensure the characters render correctly.
// Use Cases: Produce multilingual sales dashboards where chart titles must appear in Chinese and are shared as image files. | Automate report generation that embeds localized chart images in emails or web pages. | Add image‑based validation to CI pipelines by exporting charts and checking titles with OCR.
// AI Prompts: Write C# code that creates a pie chart with a Japanese title and saves it as a JPEG using Aspose.Cells. | Provide a C# method that uses an OCR library to confirm a specific Unicode string exists in a PNG exported from Aspose.Cells. | Explain how to set up font fallback in Aspose.Cells so Chinese characters render correctly in exported chart images.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTitleChineseVerification
{
    // Creates a workbook, fills cells with Chinese labels, adds a column chart, sets the title to "销售报告", and exports the chart as a PNG image. The PNG can be inspected manually or with OCR to confirm the Chinese characters are rendered correctly.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
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

            // Set the chart title using Chinese characters
            chart.Title.Text = "销售报告"; // "Sales Report" in Chinese
            chart.Title.IsVisible = true;

            // Export the chart to a PNG image
            string imagePath = "ChartWithChineseTitle.png";
            chart.ToImage(imagePath, ImageType.Png);

            // Save the workbook (optional, for further inspection)
            workbook.Save("ChartWithChineseTitle.xlsx");

            // At this point, the PNG file "ChartWithChineseTitle.png" contains the chart
            // with the title rendered in Chinese characters. Manual visual verification
            // or OCR can be used to confirm the presence of the Chinese title.
            Console.WriteLine("Chart exported to PNG with Chinese title.");
        }
    }
}
