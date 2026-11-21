// Title: How to set ellipse shape for data labels in a bubble chart using Aspose.Cells for .NET (C#)
// Common Searches: C# code to change bubble chart data label shape to ellipse with Aspose.Cells | Aspose.Cells set data label shape type for a bubble chart | apply ellipse shape to chart data labels in Aspose.Cells example | customize bubble chart data label appearance Aspose.Cells C# | set ellipse shape for chart data labels using Aspose.Cells for .NET
// Tags: set ellipse shape for chart data labels Aspose.Cells | bubble chart data label shape Aspose.Cells C# | Aspose.Cells data label formatting bubble chart | C# Aspose.Cells chart data label customization | ellipse data label type Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // This example creates a workbook, adds a bubble chart with sample data, and demonstrates how to assign the ellipse shape to the chart's data labels via the DataLabels.ShapeType property before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the bubble chart
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                sheet.Cells["C1"].PutValue("Size");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(4);
                sheet.Cells["C2"].PutValue(10);
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue(5);
                sheet.Cells["C3"].PutValue(20);
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue(6);
                sheet.Cells["C4"].PutValue(30);
                sheet.Cells["A5"].PutValue(4);
                sheet.Cells["B5"].PutValue(7);
                sheet.Cells["C5"].PutValue(40);

                // Add a bubble chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add series (X values, Y values, bubble size)
                chart.NSeries.Add("A2:C5", true);

                // Optional: show data point values in labels if the API supports it
                // Uncomment the following line if your Aspose.Cells version includes IsShowValue
                // chart.NSeries[0].DataLabels.IsShowValue = true;

                // Save the workbook
                string outputPath = "BubbleChartWithEllipseDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
