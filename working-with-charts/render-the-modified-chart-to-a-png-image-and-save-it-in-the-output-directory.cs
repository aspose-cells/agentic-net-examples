// Title: How to render an Aspose.Cells column chart as a PNG image and save it using C#
// AI Prompts: Write C# code that creates a workbook, adds a column chart, sets its data range and title, then uses Chart.ToImage to generate a PNG file in a specified folder. | Show the complete steps to export an Aspose.Cells chart to a PNG image while also persisting the workbook to disk in a .NET application. | Provide a concise C# example that demonstrates using the Aspose.Cells Chart.ToImage method to render a chart as a PNG file.
// Common Searches: C# Aspose.Cells export chart to PNG file example | How to save an Aspose.Cells chart as an image in .NET | Aspose.Cells ToImage method usage for column charts | Render Excel chart to PNG with Aspose.Cells C# | Save chart image and workbook to output directory using Aspose.Cells
// Tags: Aspose.Cells export chart to PNG | C# Aspose.Cells column chart rendering | Aspose.Cells Chart.ToImage usage | save Aspose.Cells chart as image file | output directory handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The program creates a workbook, fills it with sample data, adds a column chart with a title, renders the chart to a PNG image saved in an output folder, and also saves the workbook as an .xlsx file.
class RenderChartToPng
{
    static void Main()
    {
        // Define output folder and ensure it exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Optional: modify chart appearance (e.g., title)
        chart.Title.Text = "Fruit Sales";

        // Build the full file path for the PNG image
        string imagePath = Path.Combine(outputDir, "FruitSalesChart.png");

        // Render the chart to a PNG image file
        chart.ToImage(imagePath, ImageType.Png);

        // Optionally, save the workbook for reference
        string workbookPath = Path.Combine(outputDir, "FruitSalesWorkbook.xlsx");
        workbook.Save(workbookPath);

        Console.WriteLine($"Chart image saved to: {imagePath}");
        Console.WriteLine($"Workbook saved to: {workbookPath}");
    }
}
