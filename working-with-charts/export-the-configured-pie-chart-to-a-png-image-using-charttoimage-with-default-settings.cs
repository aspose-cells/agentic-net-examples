// Title: Export a Pie Chart to PNG with Aspose.Cells Chart.ToImage in C# (default settings)
// Description: This C# example creates a workbook, fills cells A1:B4 with category and value data, adds a pie chart, sets the data range (including headers), and calls Chart.ToImage("PieChart.png"). The .png extension automatically selects the PNG format, producing a chart image without extra configuration. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells | Chart.ToImage | export chart PNG | C# pie chart image | Excel chart to image | Aspose.Cells for .NET | default image format | pie chart snapshot | GitHub sample | Aspose.Cells chart export
// Common Searches: Aspose.Cells export pie chart to PNG C# | Chart.ToImage default format example | Save Excel chart as PNG using Aspose.Cells | C# code to convert chart to image Aspose | How to generate chart image without specifying format Aspose.Cells
// Developer Intent: Create a PNG image of a pie chart directly from an Excel workbook using Aspose.Cells.
// Use Cases: Generate chart thumbnails for dashboards or reports. | Attach chart images to automated email summaries. | Serve chart graphics in web applications without exposing the Excel file.
// AI Prompts: Show how to export an Aspose.Cells chart to JPEG with custom dimensions. | Provide code to loop through all charts in a workbook and save each as a separate PNG file. | Explain how to adjust image resolution when using Chart.ToImage in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsPieChartExport
{
    // This C# example creates a workbook, fills cells A1:B4 with category and value data, adds a pie chart, sets the data range (including headers), and calls Chart.ToImage("PieChart.png"). The .png extension automatically selects the PNG format, producing a chart image without extra configuration. The workbook is then saved as an Excel file.
    class Program
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
            // The file extension determines the image format
            pieChart.ToImage("PieChart.png");

            // Optionally save the workbook (not required for the image export)
            workbook.Save("PieChartWorkbook.xlsx");

            Console.WriteLine("Pie chart exported to PieChart.png successfully.");
        }
    }
}
