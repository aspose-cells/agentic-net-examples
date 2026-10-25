// Title: Save an Aspose.Cells workbook containing a stacked progress‑bar chart to XLSX while retaining formula cells (C#)
// Common Searches: Aspose.Cells C# save workbook to XLSX with formulas preserved | create stacked bar progress chart in Excel using Aspose.Cells .NET | keep cell formulas when exporting an Aspose.Cells workbook that includes a chart | export Excel file with a progress‑bar chart and formula‑based data using Aspose.Cells | Aspose.Cells chart series from formula range and save as .xlsx
// Tags: preserve formulas on workbook save Aspose.Cells | stacked bar progress chart Aspose.Cells C# | export workbook with chart to XLSX using Aspose.Cells | formula‑driven series in Aspose.Cells chart | save Excel file with chart and formulas .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The example creates a new workbook, fills task data with static completed values and formulas for remaining percentages, adds a stacked bar chart that acts as a progress bar, and saves the workbook as an XLSX file while ensuring all formulas remain intact.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate data -----
            // Header row
            cells["A1"].PutValue("Task");
            cells["B1"].PutValue("Completed");
            cells["C1"].PutValue("Remaining");

            // Task names
            cells["A2"].PutValue("Task 1");
            cells["A3"].PutValue("Task 2");
            cells["A4"].PutValue("Task 3");

            // Completed percentages (static values)
            cells["B2"].PutValue(30);
            cells["B3"].PutValue(70);
            cells["B4"].PutValue(50);

            // Remaining percentages calculated with formulas (preserved on save)
            cells["C2"].Formula = "100-B2";
            cells["C3"].Formula = "100-B3";
            cells["C4"].Formula = "100-B4";

            // ----- Add a stacked bar chart to act as a progress bar -----
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIdx = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.Title.Text = "Progress Bar";

            // Series for Completed values
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Completed";

            // Series for Remaining values (calculated via formulas)
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Remaining";

            // Optional: hide Remaining series from legend (if supported)
            // chart.NSeries[1].IsShowInLegend = false;

            // ----- Save workbook preserving formulas -----
            workbook.Save("ProgressBarChart.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
