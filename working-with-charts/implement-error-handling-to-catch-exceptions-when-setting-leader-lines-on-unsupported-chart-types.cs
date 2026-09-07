// Title: Implement try‑catch error handling for leader line settings on unsupported chart types in Aspose.Cells for .NET
// AI Prompts: Write C# code that attempts to set HasLeaderLines and related leader line properties on a chart, and wraps the assignments in a try‑catch block to handle the exception thrown for unsupported chart types using Aspose.Cells. | Create a sample that first checks a chart’s ChartType, applies leader line styling only to supported types (e.g., pie chart), and logs a descriptive message when the operation is invalid for the current chart.
// Common Searches: Aspose.Cells catch exception when enabling leader lines on column chart C# | how to handle unsupported chart type error for leader lines in Aspose.Cells | C# Aspose.Cells check chart type before setting HasLeaderLines | leader line properties cause exception on non‑pie charts Aspose.Cells | safe way to configure leader lines for charts using Aspose.Cells .NET
// Tags: Aspose.Cells chart leader lines exception handling | Aspose.Cells set leader lines on pie chart | Aspose.Cells unsupported chart type error handling | C# Aspose.Cells chart series leader line configuration | Aspose.Cells try-catch leader line properties | Aspose.Cells chart type validation for leader lines

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a workbook, adds a column chart (which does not support leader lines) and attempts to configure leader line properties inside a try‑catch block to capture the unsupported‑type exception. It then adds a pie chart (which supports leader lines) and applies custom leader line styling before saving the file as LeaderLinesExceptionDemo.xlsx.
class LeaderLinesExceptionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // ------------------------------------------------------------
        // Example 1: Column chart (does NOT support leader lines)
        // ------------------------------------------------------------
        int columnChartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart columnChart = sheet.Charts[columnChartIdx];
        columnChart.NSeries.Add("B2:B4", true);
        columnChart.NSeries.CategoryData = "A2:A4";

        // Attempt to configure leader lines and catch any exception
        try
        {
            Series columnSeries = columnChart.NSeries[0];
            columnSeries.HasLeaderLines = true;               // May throw for unsupported chart type
            columnSeries.LeaderLines.IsAuto = false;
            columnSeries.LeaderLines.Style = LineType.Dot;
            columnSeries.LeaderLines.WeightPt = 1.0;
            columnSeries.LeaderLines.Color = Color.Red;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Leader lines are not supported for Column chart: " + ex.Message);
        }

        // ------------------------------------------------------------
        // Example 2: Pie chart (supports leader lines)
        // ------------------------------------------------------------
        int pieChartIdx = sheet.Charts.Add(ChartType.Pie, 25, 0, 40, 8);
        Chart pieChart = sheet.Charts[pieChartIdx];
        pieChart.NSeries.Add("B2:B4", true);
        pieChart.NSeries.CategoryData = "A2:A4";

        // Configure leader lines on a supported chart type
        Series pieSeries = pieChart.NSeries[0];
        pieSeries.HasLeaderLines = true;
        pieSeries.LeaderLines.IsAuto = false;
        pieSeries.LeaderLines.Style = LineType.Dot;
        pieSeries.LeaderLines.WeightPt = 1.5;
        pieSeries.LeaderLines.Color = Color.Blue;

        // Save the workbook
        workbook.Save("LeaderLinesExceptionDemo.xlsx");
    }
}
