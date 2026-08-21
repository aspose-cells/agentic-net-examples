// Title: Export Pie Chart to PNG Using Aspose.Cells Chart.ToImage in C#
// Description: Creates a workbook, fills cells A1:B4 with categories and values, adds a pie chart, sets its data range (including headers), and saves the chart as a PNG file with Chart.ToImage default settings. The workbook is also saved for reference.
// Keywords: Aspose.Cells | C# | .NET | Chart.ToImage | export chart to PNG | pie chart image | Excel chart to image | Aspose.Cells example | save chart as PNG | default export settings | Aspose.Cells API
// Common Searches: Aspose.Cells export pie chart to PNG C# | Chart.ToImage default settings example | How to save Excel chart as PNG using Aspose.Cells | C# Aspose.Cells chart image export | Export chart without options Aspose.Cells
// Developer Intent: Export the configured pie chart to a PNG image using Chart.ToImage with default settings.
// Use Cases: Generate PNG snapshots of pie charts for web dashboards | Attach chart images to automated email reports | Batch extract charts from multiple workbooks as PNG files for documentation
// AI Prompts: Write C# code to export a line chart to JPEG using Aspose.Cells Chart.ToImage with default options. | Show how to iterate through all charts in a workbook and save each as a PNG file using Aspose.Cells. | Explain how to customize image resolution and dimensions when exporting charts with Chart.ToImage in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, fills cells A1:B4 with categories and values, adds a pie chart, sets its data range (including headers), and saves the chart as a PNG file with Chart.ToImage default settings. The workbook is also saved for reference.
class ExportPieChartToPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["B3"].PutValue(800);
        sheet.Cells["B4"].PutValue(1500);

        // Add a pie chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart pieChart = sheet.Charts[chartIndex];

        // Set the data range for the chart (including headers)
        pieChart.SetChartDataRange("A1:B4", true);

        // Export the chart to a PNG image using default settings
        // The file extension determines the image format (PNG)
        pieChart.ToImage("PieChart.png", ImageType.Png);

        // Optionally, save the workbook for reference
        workbook.Save("WorkbookWithPieChart.xlsx");

        Console.WriteLine("Pie chart exported to PieChart.png successfully.");
    }
}
