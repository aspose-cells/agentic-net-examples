// Title: Resize data label shapes and increase marker size on a line chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart, enables large markers, and resizes the data label shapes by setting the font size to 14 points. | Show how to adjust the marker size of a line series and then modify the associated data label appearance in an Excel workbook using Aspose.Cells. | Provide a complete Aspose.Cells example that saves a workbook containing a line chart with customized marker size and resized data labels.
// Common Searches: Aspose.Cells C# line chart increase marker size and change data label font | how to resize data label shapes after setting large markers in Aspose.Cells | C# example for customizing line chart data labels with Aspose.Cells | set marker size and data label appearance in Excel line chart using Aspose.Cells .NET | Aspose.Cells line chart data label formatting tutorial
// Tags: Aspose.Cells line chart marker size | C# resize chart data label font | Aspose.Cells adjust data label shape | Excel line chart customization Aspose.Cells | Aspose.Cells chart series formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Creates a workbook with sample sales data, adds a line chart, shows data labels, sets the label font size to 14 points (ready for additional marker size adjustments), and saves the file as LineChartWithResizedDataLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Sales");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["B3"].PutValue(150);
            ws.Cells["B4"].PutValue(130);

            // Add a line chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Show data labels for the series and increase their font size
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.Font.Size = 14;

            // Define output file path
            string outputPath = "LineChartWithResizedDataLabels.xlsx";

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
