// Title: Aspose.Cells for .NET: Create a Combo Column‑Line Chart with a Secondary Axis (C#)
// Description: Demonstrates how to build a workbook, insert sales and profit data, add a combo chart (column for sales, line for profit), switch the profit series to a line type, plot that series on a secondary Y‑axis, set a custom axis title, and save the file as an Excel workbook.
// Keywords: Aspose.Cells combo chart secondary axis | C# Aspose.Cells line series secondary Y axis | column and line chart Aspose.Cells .NET | custom secondary axis title Aspose.Cells | Excel combo chart with two axes
// Common Searches: Aspose.Cells set secondary axis for line series | C# create combo column line chart Aspose.Cells | how to plot line series on secondary Y axis in Excel using Aspose | Aspose.Cells secondary value axis example
// Developer Intent: Add a combo chart to a worksheet and display the line series on a secondary Y‑axis for clearer comparison of different data scales.
// Use Cases: Compare monthly sales (columns) with profit margins (line) when the values have different ranges. | Generate financial reports that show revenue and growth rate side‑by‑side using separate axes. | Build a dashboard workbook where a KPI line is plotted on a secondary axis next to primary column metrics.
// AI Prompts: Show code to add a third series to the combo chart and keep it on the primary axis. | Explain how to format the secondary axis with a currency number format in Aspose.Cells. | Provide examples of customizing marker style and line color for the secondary‑axis line series.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to build a workbook, insert sales and profit data, add a combo chart (column for sales, line for profit), switch the profit series to a line type, plot that series on a secondary Y‑axis, set a custom axis title, and save the file as an Excel workbook.
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
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(55);

        // Add a combo chart (Column + Line)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // First series (column) for Sales
        chart.NSeries.Add("B2:B4", true);
        // Second series (line) for Profit
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Change the second series to a line chart type
        chart.NSeries[1].Type = ChartType.Line;

        // Plot the line series on the secondary Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize the secondary axis title
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Profit (Secondary Axis)";

        // Save the workbook
        workbook.Save("ComboChartSecondaryAxis.xlsx");
    }
}
