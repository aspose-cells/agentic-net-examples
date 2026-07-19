// Title: Set Chart Background to Light Gray and Remove Fill Pattern with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart, assigns data ranges, then changes the ChartArea and PlotArea background to LightGray and disables any fill pattern before saving the file.
// Keywords: Aspose.Cells chart background color C# | remove chart fill pattern Aspose.Cells | light gray chart area .NET | customize chart appearance Aspose.Cells | ChartArea FillFormat Pattern.None
// Common Searches: Aspose.Cells change chart area background color to light gray | C# remove default fill pattern from chart area Aspose.Cells | set plot area background color and clear pattern using Aspose.Cells | how to style chart background in Aspose.Cells for .NET
// Developer Intent: Programmatically apply a light‑gray background to a chart and eliminate any existing fill patterns in both the chart area and plot area.
// Use Cases: Standardize chart styling across reports with a uniform light‑gray background. | Ensure charts comply with corporate branding by removing default fill patterns. | Create clean, distraction‑free visualizations by applying the same background settings to chart and plot areas.
// AI Prompts: Show a C# example that sets a chart's background to light gray and disables fill patterns using Aspose.Cells. | How can I apply a LightGray background to both ChartArea and PlotArea and clear their fill patterns in Aspose.Cells for .NET? | Explain the steps to customize chart appearance—background color and fill format—when generating Excel files with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBackground
{
    // Creates a workbook, adds a column chart, assigns data ranges, then changes the ChartArea and PlotArea background to LightGray and disables any fill pattern before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
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

            // Set chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set chart area background color to light gray
            chart.ChartArea.Area.BackgroundColor = Color.LightGray;

            // Remove any default fill pattern from the chart area
            chart.ChartArea.Area.FillFormat.Pattern = FillPattern.None;

            // Optionally, also clear fill pattern for the plot area
            chart.PlotArea.Area.FillFormat.Pattern = FillPattern.None;
            chart.PlotArea.Area.BackgroundColor = Color.LightGray;

            // Save the workbook
            workbook.Save("ChartBackgroundLightGray.xlsx");
        }
    }
}
