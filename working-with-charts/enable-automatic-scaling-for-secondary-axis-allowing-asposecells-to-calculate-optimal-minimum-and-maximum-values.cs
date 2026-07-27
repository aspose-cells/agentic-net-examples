// Title: C# – Enable Automatic Scaling on the Secondary Axis of an Aspose.Cells Column Chart
// Description: Creates a workbook, adds primary and secondary data series, builds a column chart, plots the second series on the secondary value axis, and activates automatic minimum and maximum calculation by setting IsAutomaticMinValue and IsAutomaticMaxValue to true before saving the file.
// Keywords: Aspose.Cells | C# | secondary axis | automatic scaling | IsAutomaticMinValue | IsAutomaticMaxValue | Chart.SecondValueAxis | PlotOnSecondAxis | dual‑axis column chart | Excel chart axis limits
// Common Searches: Aspose.Cells set secondary axis auto min max C# | automatic scaling for secondary value axis Aspose.Cells | Chart.SecondValueAxis IsAutomaticMinValue example | dual axis chart Aspose.Cells C# auto scale | how to enable auto scaling on secondary axis in Aspose.Cells
// Developer Intent: Let Aspose.Cells compute optimal min and max values for a chart's secondary axis automatically.
// Use Cases: Generate dual‑axis column charts where the secondary axis adapts to large data ranges without manual limits. | Produce financial or scientific reports that automatically adjust axis scales when source values change. | Export Excel files with self‑scaling secondary axes to simplify maintenance of dynamic dashboards.
// AI Prompts: Show C# code that enables automatic scaling on the secondary axis of an Aspose.Cells chart. | Provide an Aspose.Cells example that creates a column chart, plots a series on the secondary axis, and sets IsAutomaticMinValue/IsAutomaticMaxValue to true. | Explain how Aspose.Cells determines the optimal minimum and maximum for a secondary axis when automatic scaling is turned on.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds primary and secondary data series, builds a column chart, plots the second series on the secondary value axis, and activates automatic minimum and maximum calculation by setting IsAutomaticMinValue and IsAutomaticMaxValue to true before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Primary");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["B4"].PutValue(300);

        worksheet.Cells["C1"].PutValue("Secondary");
        worksheet.Cells["C2"].PutValue(5000);
        worksheet.Cells["C3"].PutValue(3000);
        worksheet.Cells["C4"].PutValue(1000);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series: first uses primary axis, second uses secondary axis
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Enable automatic scaling for the secondary value axis
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.IsAutomaticMinValue = true; // Let Aspose.Cells calculate optimal minimum
        secondaryAxis.IsAutomaticMaxValue = true; // Let Aspose.Cells calculate optimal maximum
        secondaryAxis.Title.Text = "Secondary Axis";

        // Save the workbook
        workbook.Save("SecondaryAxisAutoScaling.xlsx");
    }
}
