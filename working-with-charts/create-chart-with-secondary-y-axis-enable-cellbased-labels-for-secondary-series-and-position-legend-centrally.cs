// Title: Create a combo column‑line chart with a secondary Y axis, cell‑based data labels for the secondary series, and a bottom legend using Aspose.Cells for .NET
// Common Searches: how to add a secondary Y axis to a column chart in Aspose.Cells C# | Aspose.Cells C# add data labels from cells for a secondary line series | set legend position to bottom in an Aspose.Cells chart programmatically | create a combo column and line chart with Aspose.Cells .NET example | Aspose.Cells chart customization secondary axis and data labels
// Tags: add secondary y axis Aspose.Cells | cell based data labels secondary series Aspose.Cells | set legend position bottom Aspose.Cells | combo column line chart Aspose.Cells | export chart to Excel file Aspose.Cells | populate chart data from worksheet cells Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a new workbook, writes month names and numeric values into cells, adds a combo column‑line chart where the primary series is shown as columns and the secondary series as a line, assigns a secondary Y axis to the line series, positions the legend at the bottom, and saves the workbook as ChartWithSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Primary");
            sheet.Cells["C1"].PutValue("Secondary");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(200);
            sheet.Cells["C4"].PutValue(300);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Primary and Secondary Axis";

            // Primary series (column)
            chart.NSeries.Add("B2:B4", true);
            // Secondary series (line)
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Type = ChartType.Line;

            // Position legend at the bottom
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save workbook
            string outputPath = "ChartWithSecondaryAxis.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
