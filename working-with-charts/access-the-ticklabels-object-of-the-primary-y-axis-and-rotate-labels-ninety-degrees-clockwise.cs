// Title: Rotate primary Y‑axis tick labels 90° clockwise in an Aspose.Cells chart with C#
// Common Searches: Aspose.Cells C# rotate primary Y axis tick labels 90 degrees | set rotation angle for chart value axis tick labels using Aspose.Cells .NET | how to change orientation of Y‑axis labels in an Aspose.Cells column chart | C# code to rotate chart axis tick labels clockwise with Aspose.Cells | adjust TickLabels.RotationAngle for primary value axis in Aspose.Cells workbook
// Tags: y axis tick label rotation Aspose.Cells | TickLabels.RotationAngle C# | primary value axis label orientation .NET | Aspose.Cells chart axis label rotation | column chart y axis label angle Aspose | modify chart tick label angle C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart, accesses the chart's primary Y (value) axis, retrieves its TickLabels object, sets RotationAngle to 90 degrees to rotate the labels clockwise, and saves the file as RotatedYAxisTickLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet.
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // (Optional) Add some sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the primary Y axis (value axis) of the chart.
            Axis primaryValueAxis = chart.ValueAxis;

            // Access the TickLabels object of the primary Y axis.
            TickLabels tickLabels = primaryValueAxis.TickLabels;

            // Rotate the tick labels 90 degrees clockwise.
            tickLabels.RotationAngle = 90;

            // Define output file path.
            string outputPath = "RotatedYAxisTickLabels.xlsx";

            // Save the workbook to a file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
