// Title: Define custom Y‑axis minimum and maximum values for a chart with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, disables automatic scaling, sets ValueAxis.MinValue to 5 and MaxValue to 60, and saves the file as ChartWithCustomYAxis.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# chart axis | set Y axis minimum | set Y axis maximum | custom chart scaling | disable automatic axis | ValueAxis MinValue | ValueAxis MaxValue | Excel chart programming
// Common Searches: Aspose.Cells set custom Y axis range | how to define chart Y‑axis minimum .NET | disable automatic Y axis scaling Aspose.Cells | set ValueAxis MinValue and MaxValue in C# | chart axis limits programmatically Excel
// Developer Intent: The developer wants to programmatically specify exact minimum and maximum values for a chart’s Y‑axis to control its visual scale.
// Use Cases: Produce reports where all charts share a fixed Y‑axis range for easy comparison. | Generate financial or scientific spreadsheets that require a baseline start point on the Y‑axis. | Create dashboards that must adhere to corporate charting standards with predefined axis limits.
// AI Prompts: Write C# code with Aspose.Cells to set the Y‑axis minimum to 0 and maximum to 100 for a line chart. | Explain how to turn off automatic axis scaling and assign custom MinValue and MaxValue for any chart type in Aspose.Cells. | Provide step‑by‑step instructions to locate and modify the ValueAxis of an existing chart in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, disables automatic scaling, sets ValueAxis.MinValue to 5 and MaxValue to 60, and saves the file as ChartWithCustomYAxis.xlsx using Aspose.Cells.
class SetYAxisScale
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(50);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the Y‑axis (value axis) and define custom scaling
        Axis valueAxis = chart.ValueAxis;
        valueAxis.IsAutomaticMinValue = false; // turn off automatic minimum
        valueAxis.MinValue = 5;                // set desired minimum value
        valueAxis.IsAutomaticMaxValue = false; // turn off automatic maximum
        valueAxis.MaxValue = 60;               // set desired maximum value

        // Save the workbook with the customized chart
        workbook.Save("ChartWithCustomYAxis.xlsx");
    }
}
