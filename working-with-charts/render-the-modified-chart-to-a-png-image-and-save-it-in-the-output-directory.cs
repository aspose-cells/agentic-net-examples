// Title: Export Aspose.Cells Chart to PNG in C# and Save to an Output Folder
// Description: C# example that creates a workbook, adds sample data, builds a column chart, sets a title, creates an "output" directory, renders the chart to a PNG file with Chart.ToImage, and saves both the image and the workbook in that folder.
// Keywords: Aspose.Cells chart export C# | Chart.ToImage PNG Aspose.Cells | save chart image to folder | export Excel chart as PNG | .NET chart to image example | Aspose.Cells render chart PNG | C# Aspose.Cells image export
// Common Searches: How to export an Aspose.Cells chart to PNG using C# | Aspose.Cells Chart.ToImage example code | Save chart image to specific directory Aspose.Cells | Render Excel column chart as PNG in .NET | Aspose.Cells export chart image tutorial
// Developer Intent: Render a modified Aspose.Cells chart as a PNG file and store it in a designated output directory.
// Use Cases: Generate PNG snapshots of Excel charts for automated reporting pipelines. | Create image assets of charts for embedding in web pages, emails, or mobile apps without sharing the workbook. | Batch‑export multiple charts from a workbook to PNG files for archival or publishing purposes.
// AI Prompts: Show how to export an Aspose.Cells chart to JPEG instead of PNG in C#. | Provide code that iterates through all charts in a workbook and saves each as a separate PNG file. | Explain how to adjust image resolution or DPI when using Chart.ToImage in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds sample data, builds a column chart, sets a title, creates an "output" directory, renders the chart to a PNG file with Chart.ToImage, and saves both the image and the workbook in that folder.
class Program
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

        // Modify the chart (e.g., set a title)
        chart.Title.Text = "Fruit Sales";

        // Ensure the output directory exists
        string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
        Directory.CreateDirectory(outputDir);

        // Render the chart to a PNG image and save it using the ToImage(string, ImageType) overload
        string imagePath = Path.Combine(outputDir, "FruitSalesChart.png");
        chart.ToImage(imagePath, ImageType.Png);

        // Optionally save the workbook that contains the chart
        workbook.Save(Path.Combine(outputDir, "WorkbookWithChart.xlsx"));
    }
}
