// Title: Aspose.Cells C# – Create a Combo Chart with Column, Line, and Area Series on Separate Axes
// Description: This example builds a new workbook, fills A2:D4 with monthly data, adds a combo chart, defines three series (column, line, area), assigns the line series to the secondary Y‑axis, and saves the file as ComboChart.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | combo chart | column series | line series | area series | secondary axis | PlotOnSecondAxis | Excel chart types | multi‑type chart | chart customization
// Common Searches: Aspose.Cells combo chart C# example | how to add column, line and area series in one chart | set secondary Y axis for a series Aspose.Cells | create Excel combo chart with multiple chart types .NET | PlotOnSecondAxis Aspose.Cells tutorial
// Developer Intent: Generate an Excel workbook that contains a combo chart combining column, line, and area series, with the line series plotted on a secondary Y‑axis.
// Use Cases: Show sales volume (column), profit margin (line) and cumulative revenue (area) together for a financial dashboard. | Display temperature (area), precipitation (column) and wind speed (line) on separate axes in a weather report. | Monitor manufacturing performance by charting production count (column), defect rate (line) and cumulative output (area).
// AI Prompts: Write C# code with Aspose.Cells to add a combo chart that includes column, line, and area series, where the line series uses the secondary Y‑axis. | Explain how PlotOnSecondAxis works in an Aspose.Cells combo chart and how it influences axis rendering. | Provide steps to customize titles, legends, and axis labels for a multi‑type combo chart created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example builds a new workbook, fills A2:D4 with monthly data, adds a combo chart, defines three series (column, line, area), assigns the line series to the secondary Y‑axis, and saves the file as ComboChart.xlsx using Aspose.Cells for .NET.
class ComboChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Categories
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        // Column series values
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Line series values
        sheet.Cells["C2"].PutValue(80);
        sheet.Cells["C3"].PutValue(130);
        sheet.Cells["C4"].PutValue(170);

        // Area series values
        sheet.Cells["D2"].PutValue(60);
        sheet.Cells["D3"].PutValue(110);
        sheet.Cells["D4"].PutValue(140);

        // Add a combo chart (initially as Column type)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.Title.Text = "Combo Chart: Column, Line, Area";

        // Set the category (X) axis data for all series
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

        // ----- Series 1: Column (primary axis) -----
        int colSeriesIdx = chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
        Series colSeries = chart.NSeries[colSeriesIdx];
        colSeries.Type = ChartType.Column;
        colSeries.Name = "Column Series";
        // Plotted on primary value axis by default

        // ----- Series 2: Line (secondary axis) -----
        int lineSeriesIdx = chart.NSeries.Add("=Sheet1!$C$2:$C$4", true);
        Series lineSeries = chart.NSeries[lineSeriesIdx];
        lineSeries.Type = ChartType.Line;
        lineSeries.Name = "Line Series";
        lineSeries.PlotOnSecondAxis = true; // Use secondary Y axis

        // ----- Series 3: Area (uses series axis) -----
        int areaSeriesIdx = chart.NSeries.Add("=Sheet1!$D$2:$D$4", true);
        Series areaSeries = chart.NSeries[areaSeriesIdx];
        areaSeries.Type = ChartType.Area;
        areaSeries.Name = "Area Series";
        // Area series will be plotted on the primary axis; if a distinct axis is required,
        // you can also set PlotOnSecondAxis = true. Here we keep it on the primary axis.

        // Save the workbook with the combo chart
        workbook.Save("ComboChart.xlsx");
    }
}
