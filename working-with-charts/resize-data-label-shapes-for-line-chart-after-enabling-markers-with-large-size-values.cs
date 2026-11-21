// Title: Resize data label shapes for a line chart with large markers using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# resize data label shape after enabling large markers in a line chart | how to change chart data label dimensions for line charts with big markers in Aspose.Cells | C# Aspose.Cells adjust data label size when marker size is increased on a line chart | set custom size for line chart data labels after applying large markers with Aspose.Cells | increase data label shape width and height for line chart markers using Aspose.Cells .NET
// Tags: resize chart data label shape Aspose.Cells | line chart data label dimensions .NET | adjust data label size after large marker Aspose.Cells | customize data label shape for line chart Aspose.Cells | Aspose.Cells set data label size for markers

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample sales data, inserts a line chart, enables data labels, and demonstrates how to resize the data label shapes when large markers are used, then saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the line chart
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Sales");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["B3"].PutValue(150);
            ws.Cells["B4"].PutValue(130);

            // Add a line chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);

            // Optional: enable data labels to show values
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Save the workbook
            string outputPath = "LineChartWithResizedDataLabels.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
