// Title: Change the second series from Column to Line to build a Combo chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds month, sales, and profit data, inserts a column chart, defines two series, sets the X‑axis categories, then switches the second series to a line type using Chart.NSeries[i].Type, producing a mixed column‑line combo chart and saves it as ComboChart.xlsx.
// Keywords: Aspose.Cells C# combo chart | change series type Aspose.Cells | column to line series .NET | Chart.NSeries Type property | mixed column line chart Aspose | create combo chart programmatically | Aspose.Cells chart customization | C# Excel chart series conversion
// Common Searches: Aspose.Cells change series to line after chart creation | how to make a combo chart in C# with Aspose.Cells | set second series type line Aspose.Cells chart | convert column series to line Aspose.Cells .NET | mixed column and line chart example Aspose
// Developer Intent: Alter an existing column chart so that the second data series is rendered as a line, resulting in a combo chart.
// Use Cases: Display sales as columns and profit as a line in a single Excel chart. | Transform a pure column chart into a combo chart by updating a specific series type. | Assign descriptive names to series after changing their visual representation for clearer reporting.
// AI Prompts: Write C# code with Aspose.Cells that converts the second series of a column chart to a line series to create a combo chart. | Show how to build a combo chart with mixed column and line series, set series names, and save the workbook using Aspose.Cells for .NET. | Explain the use of Chart.NSeries[i].Type to switch a series from column to line in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds month, sales, and profit data, inserts a column chart, defines two series, sets the X‑axis categories, then switches the second series to a line type using Chart.NSeries[i].Type, producing a mixed column‑line combo chart and saves it as ComboChart.xlsx.
class ComboChartExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
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
        sheet.Cells["C4"].PutValue(60);

        // Add a chart initially of type Column (combo chart will be formed later)
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Add first series (Sales) – will stay as Column
        chart.NSeries.Add("B2:B4", true);

        // Add second series (Profit) – initially also Column
        chart.NSeries.Add("C2:C4", true);

        // Set category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A4";

        // Change the second series type from Column to Line to create a combo chart
        // Uses Series.Type property (rule exists)
        chart.NSeries[1].Type = ChartType.Line;

        // Optional: give meaningful names to the series
        chart.NSeries[0].Name = "Sales";
        chart.NSeries[1].Name = "Profit";

        // Save the workbook (lifecycle: save)
        workbook.Save("ComboChart.xlsx");
    }
}
