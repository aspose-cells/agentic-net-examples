// Title: Export an Aspose.Cells chart to PNG in C# and save it to an output folder
// Description: Creates a workbook, fills it with sample data, adds a column chart, ensures an "output" directory exists, and uses Aspose.Cells' `Chart.ToImage` method with `ImageType.Png` to write the chart as a PNG file. The workbook can also be saved for reference.
// Keywords: Aspose.Cells chart export PNG | C# Aspose.Cells ToImage | save Excel chart as image | export chart to PNG .NET | Aspose.Cells image generation
// Common Searches: how to export Aspose.Cells chart as PNG in C# | Aspose.Cells save chart image to folder | C# export Excel chart to PNG file | Aspose.Cells ToImage example | create output directory and export chart image
// Developer Intent: Generate a PNG image from an Aspose.Cells chart and store it in a specified directory.
// Use Cases: Create visual chart assets for dashboards without sharing the Excel file. | Automate batch jobs that produce PNG snapshots of charts for reporting emails. | Prepare chart images for web pages or mobile apps where only the graphic is needed.
// AI Prompts: Show C# code to export an Aspose.Cells chart to PNG with custom width and height. | Provide a loop that saves every chart in a workbook as separate PNG files. | Explain how to adjust image quality or switch to JPEG when exporting charts with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, fills it with sample data, adds a column chart, ensures an "output" directory exists, and uses Aspose.Cells' `Chart.ToImage` method with `ImageType.Png` to write the chart as a PNG file. The workbook can also be saved for reference.
class ExportChartToPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["A4"].PutValue("Banana");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1200);
        worksheet.Cells["B3"].PutValue(800);
        worksheet.Cells["B4"].PutValue(1500);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Ensure the output folder exists
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // Export the chart as a PNG image
        string imagePath = Path.Combine(outputFolder, "ChartImage.png");
        chart.ToImage(imagePath, ImageType.Png);

        // Optionally save the workbook for reference
        workbook.Save(Path.Combine(outputFolder, "Workbook.xlsx"));
    }
}
