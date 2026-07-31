// Title: Set Light‑Gray Background for Chart and Plot Areas in Aspose.Cells (.NET)
// Description: Creates a workbook, adds sample data, inserts a column chart, and applies an opaque light‑gray background to both the ChartArea and PlotArea while removing any fill pattern, then saves the file as ChartBackgroundLightGray.xlsx.
// Keywords: Aspose.Cells | C# | chart background color | light gray chart area | remove fill pattern | ChartArea BackgroundMode Opaque | PlotArea BackgroundMode Opaque | FillPattern.None | Excel chart styling | Aspose.Cells chart customization
// Common Searches: Aspose.Cells set chart area background color | remove chart fill pattern Aspose.Cells .NET | light gray background for Excel chart using Aspose | make chart background opaque Aspose.Cells | C# code to change plot area color Aspose.Cells
// Developer Intent: Apply an opaque light‑gray, pattern‑free background to a chart’s ChartArea and PlotArea in an Aspose.Cells workbook using C#.
// Use Cases: Standardize the appearance of charts in automated reports by giving them a solid light‑gray background. | Prepare printable Excel files where chart backgrounds must be uniform and free of patterns. | Apply consistent background styling to multiple charts across a workbook with minimal code.
// AI Prompts: Generate C# code with Aspose.Cells that sets both ChartArea and PlotArea to a light‑gray opaque background and removes any fill pattern. | Show how to change the background color of an Aspose.Cells chart to light gray and disable fill patterns in .NET. | Explain the steps to make chart and plot areas solid light gray in an Excel workbook using Aspose.Cells for C#.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a column chart, and applies an opaque light‑gray background to both the ChartArea and PlotArea while removing any fill pattern, then saves the file as ChartBackgroundLightGray.xlsx.
class SetChartBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(20);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set chart area background to light gray and make it opaque
        chart.ChartArea.BackgroundMode = BackgroundMode.Opaque;
        chart.ChartArea.Area.BackgroundColor = Color.LightGray;
        // Remove any fill pattern from the chart area
        chart.ChartArea.Area.FillFormat.Pattern = FillPattern.None;

        // Also apply the same settings to the plot area (optional but ensures consistency)
        chart.PlotArea.BackgroundMode = BackgroundMode.Opaque;
        chart.PlotArea.Area.BackgroundColor = Color.LightGray;
        chart.PlotArea.Area.FillFormat.Pattern = FillPattern.None;

        // Save the workbook
        workbook.Save("ChartBackgroundLightGray.xlsx");
    }
}
