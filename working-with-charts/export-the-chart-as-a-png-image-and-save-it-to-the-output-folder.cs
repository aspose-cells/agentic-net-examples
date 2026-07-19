// Title: Export a Column Chart to PNG with Aspose.Cells for .NET (C#) and Save to an Output Folder
// Description: Creates a workbook, fills cells A1:B4 with sample sales data, adds a column chart, ensures an "output" directory exists, and uses the Chart.ToImage method to write the chart as a PNG file named "ChartImage.png" into that folder.
// Keywords: Aspose.Cells | C# | .NET | chart export PNG | Chart.ToImage | column chart | output directory | image export | Aspose.Cells example
// Common Searches: Aspose.Cells export chart to PNG C# | Save chart image to folder using Aspose.Cells | How to use Chart.ToImage in Aspose.Cells .NET | Export column chart as PNG file Aspose.Cells | Create output folder and export chart image Aspose.Cells
// Developer Intent: Generate a column chart in a workbook and write it as a PNG file to a specified output directory.
// Use Cases: Include a chart thumbnail in a web dashboard by exporting it as a PNG. | Automate report generation where charts are saved as images for PDF or email attachments. | Batch‑process workbooks to extract all charts as PNG files for documentation or archival.
// AI Prompts: Show C# code that exports an Aspose.Cells chart to JPEG instead of PNG. | Provide a loop that saves every chart in a workbook as separate PNG files. | Explain how to set custom image resolution and dimensions when exporting a chart with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, fills cells A1:B4 with sample sales data, adds a column chart, ensures an "output" directory exists, and uses the Chart.ToImage method to write the chart as a PNG file named "ChartImage.png" into that folder.
class ExportChartAsPng
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
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
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Ensure the output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Export the chart as a PNG image
        string imagePath = Path.Combine(outputDir, "ChartImage.png");
        chart.ToImage(imagePath, ImageType.Png);

        Console.WriteLine($"Chart exported successfully to: {imagePath}");
    }
}
