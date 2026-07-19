// Title: C# – Add Percentage Data Labels to a Stacked Bar Progress Chart using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert task names and progress values, generate a stacked bar (progress) chart, and configure data labels to show exact percentages (0.00%) inside each bar, then save as XLSX.
// Keywords: Aspose.Cells | C# chart data labels | stacked bar progress chart | percentage labels | inside end label | Excel automation | chart formatting | ShowPercentage Aspose.Cells
// Common Searches: Aspose.Cells show percentage on stacked bar chart | C# add data labels to chart series | progress bar chart with percentages Aspose.Cells | format chart data labels as 0.00% in .NET | position data labels inside bar Aspose.Cells
// Developer Intent: Add data labels that display exact percentage values to the visible series of a stacked bar progress chart.
// Use Cases: Build a project‑status dashboard with progress bars that include inside‑bar percentage labels. | Generate automated Excel reports where each task’s completion is visualized with a labeled progress bar. | Create templates for status‑update spreadsheets that combine visual bars and precise percentage text.
// AI Prompts: Write C# code with Aspose.Cells to add 0.00% data labels to a stacked bar chart and place them inside the bar. | Explain how to hide raw values and show only formatted percentages on chart series in Aspose.Cells for .NET. | Show how to apply a custom number format to chart data labels and ensure they appear only for visible series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartWithDataLabels
{
    // Demonstrates how to create a workbook, insert task names and progress values, generate a stacked bar (progress) chart, and configure data labels to show exact percentages (0.00%) inside each bar, then save as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and add sample data
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Cells["A1"].PutValue("Task");
                dataSheet.Cells["B1"].PutValue("Progress");

                dataSheet.Cells["A2"].PutValue("Design");
                dataSheet.Cells["B2"].PutValue(0.30); // 30%
                dataSheet.Cells["A3"].PutValue("Development");
                dataSheet.Cells["B3"].PutValue(0.55); // 55%
                dataSheet.Cells["A4"].PutValue("Testing");
                dataSheet.Cells["B4"].PutValue(0.15); // 15%

                // Add a chart sheet for the progress bar chart
                Worksheet chartSheet = workbook.Worksheets[workbook.Worksheets.Add(SheetType.Chart)];

                // Create a stacked bar chart (commonly used for progress bars)
                int chartIndex = chartSheet.Charts.Add(ChartType.BarStacked, 0, 0, 20, 12);
                Chart chart = chartSheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure data labels for each series
                foreach (Series series in chart.NSeries)
                {
                    // Enable data labels and show only the percentage value
                    series.DataLabels.ShowPercentage = true;
                    series.DataLabels.ShowValue = false; // hide raw value
                    // Use a number format that displays exact percentages (e.g., 30.00%)
                    series.DataLabels.NumberFormat = "0.00%";
                    // Position the label inside the bar for better readability
                    series.DataLabels.Position = LabelPositionType.InsideEnd;
                }

                // Save the workbook
                workbook.Save("ProgressBarChartWithDataLabels.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
