// Title: Aspose.Cells for .NET – Add a Line Series to the Secondary Axis in a Column‑Line Combo Chart
// Description: This example creates a workbook, fills it with monthly sales and profit data, inserts a column‑line combo chart, moves the profit line series to a secondary Y‑axis, customizes the secondary axis title and scale, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells combo chart secondary axis | C# secondary Y axis line series | column line chart Aspose.Cells .NET | customize secondary value axis | Excel chart programming C#
// Common Searches: Aspose.Cells set line series on secondary axis | C# combo chart with two Y axes using Aspose.Cells | how to add secondary axis to chart in Aspose.Cells | column and line chart with separate axes .NET
// Developer Intent: Place the line series of a combo chart on a secondary Y‑axis and adjust its axis properties programmatically.
// Use Cases: Compare sales (columns) and profit margins (line) that have different value ranges in a single report. | Create financial dashboards where revenue and expense trends need distinct scales. | Provide a reusable method for adding a column‑line combo chart with independent axes to any Excel export.
// AI Prompts: Write C# code to format the secondary axis number display as currency in an Aspose.Cells combo chart. | Show how to enable data labels for both primary and secondary series in a column‑line combo chart. | Explain how to switch the secondary axis to a logarithmic scale using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, fills it with monthly sales and profit data, inserts a column‑line combo chart, moves the profit line series to a secondary Y‑axis, customizes the secondary axis title and scale, and saves the file as an Excel workbook.
class ComboChartSecondaryAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(200);
        sheet.Cells["B3"].PutValue(250);
        sheet.Cells["B4"].PutValue(300);

        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(20);
        sheet.Cells["C3"].PutValue(30);
        sheet.Cells["C4"].PutValue(25);

        // Add a combo chart (Column + Line)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // First series (column) – Sales
        chart.NSeries.Add("B2:B4", true);
        // Second series (line) – Profit
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the second series to be a line chart
        chart.NSeries[1].Type = ChartType.Line;

        // Plot the line series on the secondary Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize the secondary value axis
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Profit (Secondary Axis)";
        secondaryAxis.MinValue = 0;
        secondaryAxis.MaxValue = 50;
        secondaryAxis.MajorUnit = 10;

        // Save the workbook
        workbook.Save("ComboChartSecondaryAxis.xlsx");
    }
}
