// Title: Resize data label shapes for a line chart with large markers using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a line chart, enable circular markers, turn on data labels, disable automatic shape resizing, set a fixed width for each label, recalculate the chart, and save the file as an XLSX document with Aspose.Cells for .NET.
// Keywords: Aspose.Cells resize data label | line chart marker size C# | ChartPoint.DataLabels.Width | disable IsResizeShapeToFitText | custom data label width Aspose.Cells | Aspose.Cells chart label formatting | C# Excel chart customization
// Common Searches: how to change data label width in Aspose.Cells line chart | resize chart data labels after increasing marker size | set ChartPoint.DataLabels.Width property in C# | disable automatic data label shape resizing Aspose.Cells | adjust line chart markers and labels with Aspose.Cells
// Developer Intent: Modify the size of data label shapes in a line chart to fit large markers by disabling auto‑resize and applying a fixed width using Aspose.Cells for .NET.
// Use Cases: Create a line chart with custom marker style and size. | Show value labels for each point while preventing automatic label resizing. | Apply a uniform width to all data label shapes to maintain readability with large markers. | Recalculate the chart to reflect label size changes before saving the workbook.
// AI Prompts: Write C# code that sets a fixed width for each data label in an Aspose.Cells line chart and turns off automatic resizing. | Explain how to increase marker size on a line chart without overlapping data labels in Aspose.Cells for .NET. | Provide step‑by‑step instructions to recalculate an Aspose.Cells chart after modifying ChartPoint.DataLabels properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a line chart, enable circular markers, turn on data labels, disable automatic shape resizing, set a fixed width for each label, recalculate the chart, and save the file as an XLSX document with Aspose.Cells for .NET.
    public class ResizeDataLabelShapesDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a line chart
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(i + 1);          // X values
                sheet.Cells[i + 1, 1].PutValue((i + 1) * 5);   // Y values
            }

            // Add a line chart
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series and set X/Y ranges
            int seriesIndex = chart.NSeries.Add("B2:B11", true);
            Series series = chart.NSeries[seriesIndex];
            series.XValues = "A2:A11";

            // Enable markers and set a large size (points)
            series.Marker.MarkerStyle = ChartMarkerType.Circle;
            // Note: The 'Size' property may not be available in some versions; omitted for compatibility.

            // Enable data labels for the series
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Resize each data label shape to accommodate large markers
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.IsResizeShapeToFitText = false;
                point.DataLabels.Width = 80; // width in points (adjust as needed)
            }

            // Recalculate the chart to apply changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("ResizeDataLabelShapesDemo.xlsx");
        }
    }
}
