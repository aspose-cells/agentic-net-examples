// Title: Export a Column Chart with Chinese Labels to PNG using Aspose.Cells for .NET
// Description: Creates a workbook, writes Chinese category names and numeric data, generates a column chart, renders it directly to a PNG file (LocalizedChart.png), and verifies the image file exists. Ideal for testing that Unicode characters are preserved in the exported chart image.
// Keywords: Aspose.Cells | C# chart export | Unicode chart labels | Chinese Excel chart | PNG image generation | chart to image .NET | validate chart image metadata
// Common Searches: Aspose.Cells export chart to PNG | how to add Chinese text to Excel chart with Aspose | verify Unicode in chart image Aspose.Cells | C# render chart as image | check chart image file exists
// Developer Intent: Generate a column chart with Chinese category labels and save it as a PNG image, then confirm the file was created.
// Use Cases: Display multilingual data in Excel charts for reports or dashboards. | Create image assets from spreadsheets for web or PDF integration. | Automate validation that exported chart images retain Unicode characters.
// AI Prompts: Write C# code that reads the PNG metadata of LocalizedChart.png and confirms Unicode text is present. | Show how to programmatically verify that Chinese labels appear correctly in the exported chart image. | Explain error‑handling strategies when Aspose.Cells fails to render a chart containing unsupported characters.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, writes Chinese category names and numeric data, generates a column chart, renders it directly to a PNG file (LocalizedChart.png), and verifies the image file exists. Ideal for testing that Unicode characters are preserved in the exported chart image.
class ValidateChartImageUnicode
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add Unicode (Chinese) text to cells – this will be used as chart categories
            sheet.Cells["A1"].PutValue("类别"); // "Category" in Chinese
            sheet.Cells["A2"].PutValue("苹果"); // "Apple"
            sheet.Cells["A3"].PutValue("橙子"); // "Orange"
            sheet.Cells["A4"].PutValue("香蕉"); // "Banana"

            // Add numeric data
            sheet.Cells["B1"].PutValue("数量"); // "Quantity"
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Create a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories (Unicode)

            // Define output image path
            string outputPath = "LocalizedChart.png";

            // Render the chart directly to a PNG file
            chart.ToImage(outputPath);

            // Verify that the image file was created
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Chart image created successfully.");
                Console.WriteLine($"Chart image saved to: {Path.GetFullPath(outputPath)}");
            }
            else
            {
                Console.WriteLine($"Failed to create chart image at: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
