// Title: Save Aspose.Cells Chart as PNG with a Timestamped Filename in C#
// Description: Creates an output folder, builds a simple column chart from sample data, generates a filename that includes the current date, time and milliseconds, and uses Chart.ToImage to export the chart as a PNG file to the specified path.
// Keywords: Aspose.Cells chart export PNG | C# timestamped filename | Chart.ToImage example | save chart to folder Aspose.Cells | .NET workbook chart image | dynamic chart file name
// Common Searches: Aspose.Cells export chart to PNG with unique name | C# save chart image to specific directory | timestamped file name for Aspose.Cells chart | Chart.ToImage C# Aspose.Cells example | create output folder and save chart image
// Developer Intent: Generate a column chart and write it as a PNG file to a chosen directory using a filename that contains the current timestamp.
// Use Cases: Automated reporting systems that need non‑overlapping chart images. | Batch processing of workbooks where each chart snapshot is stored with a date‑time stamp for audit trails. | Web APIs that return a path to a newly created chart image for downstream consumption.
// AI Prompts: Provide C# code that creates a line chart from worksheet data and saves it as a JPEG with a timestamped filename in a user‑specified folder using Aspose.Cells. | Write a method that accepts a Workbook and an output directory, iterates over all charts, and saves each as a PNG file named with the chart title and current timestamp.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageSaveDemo
{
    // Creates an output folder, builds a simple column chart from sample data, generates a filename that includes the current date, time and milliseconds, and uses Chart.ToImage to export the chart as a PNG file to the specified path.
    class Program
    {
        static void Main()
        {
            // Define the output directory
            string outputDir = "output";

            // Ensure the directory exists (CreateDirectory is a SaveOptions property,
            // but here we create the folder manually before saving)
            Directory.CreateDirectory(outputDir);

            // Generate a timestamped filename for the PNG image
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
            string fileName = $"Chart_{timestamp}.png";
            string filePath = Path.Combine(outputDir, fileName);

            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Save the chart as a PNG image using the timestamped filename
            chart.ToImage(filePath, ImageType.Png);

            Console.WriteLine($"Chart image saved to: {Path.GetFullPath(filePath)}");
        }
    }
}
