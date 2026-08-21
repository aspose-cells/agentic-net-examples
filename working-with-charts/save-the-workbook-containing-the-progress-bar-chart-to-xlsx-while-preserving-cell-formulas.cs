// Title: Save a Progress Bar Chart Workbook to XLSX with Formulas Using Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds task names, completed percentages and a formula‑based remaining column, builds a stacked bar chart to act as a progress bar, and saves the file as XLSX while keeping all cell formulas intact.
// Keywords: Aspose.Cells | C# | .NET | SaveFormat.Xlsx | stacked bar chart | progress bar chart | preserve formulas | workbook export | Excel automation | chart example
// Common Searches: Aspose.Cells save workbook to xlsx with formulas | C# stacked progress bar chart Aspose.Cells | export chart workbook without losing formulas | how to keep formulas when saving Aspose.Cells file | create progress bar chart in Excel using Aspose.Cells
// Developer Intent: Generate an Excel workbook that contains a stacked progress‑bar chart and save it as an XLSX file while ensuring all formulas remain functional.
// Use Cases: Automated project status reports that include editable progress‑bar visuals. | Weekly task‑tracking sheets where remaining work is calculated by formulas and displayed in a chart. | Dashboards that combine formula‑driven data with stacked bar charts for end‑user interaction in Excel.
// AI Prompts: Write C# code with Aspose.Cells to add a stacked progress‑bar chart and save the workbook as XLSX preserving formulas. | Show how to modify the chart type or data range in the example without breaking formula references. | Suggest ways to apply conditional formatting to the progress‑bar cells before exporting the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartSaveExample
{
    // C# example that creates a workbook, adds task names, completed percentages and a formula‑based remaining column, builds a stacked bar chart to act as a progress bar, and saves the file as XLSX while keeping all cell formulas intact.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Prepare data for a simple progress bar chart
                // -------------------------------------------------
                // Column A: Task names
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["A4"].PutValue("Testing");

                // Column B: Completed percentage (as numbers)
                sheet.Cells["B1"].PutValue("Completed");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(60);
                sheet.Cells["B4"].PutValue(90);

                // Column C: Remaining percentage calculated by a formula
                sheet.Cells["C1"].PutValue("Remaining");
                // Formula: =100-B2 (and copy down)
                sheet.Cells["C2"].Formula = "=100-B2";
                sheet.Cells["C3"].Formula = "=100-B3";
                sheet.Cells["C4"].Formula = "=100-B4";

                // -------------------------------------------------
                // Add a stacked bar chart to represent progress
                // -------------------------------------------------
                // Use BarStacked (correct enum value) for a stacked bar chart
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add Completed series (first part of the bar)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Completed";

                // Add Remaining series (second part of the bar)
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Remaining";

                // Set category (task names)
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: format the chart to look like a progress bar
                chart.Title.Text = "Project Progress";
                chart.Legend.Position = LegendPositionType.Bottom;

                // -------------------------------------------------
                // Save the workbook to XLSX while preserving formulas
                // -------------------------------------------------
                string outputPath = "ProgressBarChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
