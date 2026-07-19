// Title: Export a Modified Aspose.Cells Chart to PNG in C# and Save to Output Folder
// Description: This example creates a workbook, fills it with fruit‑sales data, adds a column chart, sets a title, and uses the Chart.ToImage method to render the chart as a PNG file stored in an output directory. The workbook can also be saved as an XLSX file for further inspection.
// Keywords: Aspose.Cells | C# chart export PNG | Chart.ToImage | save chart image | export Aspose chart as PNG | column chart image | output folder | Aspose.Cells example | render chart to image
// Common Searches: How to export an Aspose.Cells chart to PNG in C# | Aspose.Cells Chart.ToImage usage | Save Aspose chart as image file | Render column chart to PNG with Aspose.Cells | C# export chart image to folder
// Developer Intent: Render a modified Aspose.Cells chart as a PNG file and store it in a designated output directory.
// Use Cases: Generate chart images for web dashboards or reporting portals. | Create thumbnail graphics for email summaries or notifications. | Automate batch conversion of workbook charts to PNG files for documentation.
// AI Prompts: Show how to export the chart as a JPEG instead of PNG using Aspose.Cells in C#. | Provide code to set custom dimensions for the chart image before calling Chart.ToImage. | Explain how to embed the exported PNG chart into a PDF with Aspose.PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // This example creates a workbook, fills it with fruit‑sales data, adds a column chart, sets a title, and uses the Chart.ToImage method to render the chart as a PNG file stored in an output directory. The workbook can also be saved as an XLSX file for further inspection.
    class Program
    {
        static void Main()
        {
            // Define output directory and ensure it exists
            string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
            Directory.CreateDirectory(outputDir);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["B4"].PutValue(1500);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Modify chart appearance (example: set title)
            chart.Title.Text = "Fruit Sales";

            // Build full file path for the PNG image
            string imagePath = Path.Combine(outputDir, "ModifiedChart.png");

            // Render the chart to a PNG image and save it
            chart.ToImage(imagePath, ImageType.Png);

            // Optionally, save the workbook if further inspection is needed
            string workbookPath = Path.Combine(outputDir, "ChartWorkbook.xlsx");
            workbook.Save(workbookPath);

            Console.WriteLine($"Chart image saved to: {imagePath}");
            Console.WriteLine($"Workbook saved to: {workbookPath}");
        }
    }
}
