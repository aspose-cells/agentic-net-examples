// Title: C# – Assign a Line Chart Series to the Secondary Axis with Aspose.Cells
// Description: Shows how to create a workbook, add sample data, build a line chart, and move the second series to the secondary vertical axis by setting PlotOnSecondAxis (AxisGroup = 2) in Aspose.Cells for .NET, then save the file as an .xlsx workbook.
// Keywords: Aspose.Cells | C# | line chart | secondary axis | PlotOnSecondAxis | AxisGroup | Excel chart series | chart API example | Aspose.Cells tutorial
// Common Searches: Aspose.Cells secondary axis C# | plot series on secondary vertical axis Aspose.Cells | Chart.NSeries PlotOnSecondAxis example | how to use AxisGroup in Aspose.Cells | line chart with two axes .NET
// Developer Intent: Configure a chart series to render on the secondary vertical axis in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Display unit sales on the primary axis and revenue on a secondary axis to compare different scales in a single line chart. | Generate financial reports where profit margin and stock price are plotted together, each on its own vertical axis. | Create dashboards that combine temperature (primary) and humidity (secondary) readings for clearer visual analysis.
// AI Prompts: Provide a C# snippet that assigns a line chart series to the secondary axis using Aspose.Cells. | Explain the difference between PlotOnSecondAxis and AxisGroup when positioning chart series in Aspose.Cells for .NET. | Show step‑by‑step code to build a line chart with one series on the primary axis and another on the secondary axis, then save the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add sample data, build a line chart, and move the second series to the secondary vertical axis by setting PlotOnSecondAxis (AxisGroup = 2) in Aspose.Cells for .NET, then save the file as an .xlsx workbook.
class AssignSeriesToSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        sheet.Cells["B1"].PutValue("Primary");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Secondary");
        sheet.Cells["C2"].PutValue(100);
        sheet.Cells["C3"].PutValue(200);
        sheet.Cells["C4"].PutValue(300);

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: first for primary axis, second for secondary axis
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Assign the second series to the secondary vertical axis (AxisGroup = 2)
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Save the workbook
        workbook.Save("LineSeriesSecondaryAxis.xlsx");
    }
}
