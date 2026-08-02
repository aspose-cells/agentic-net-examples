// Title: C# – Add Percentage Data Labels to a Stacked Bar Progress Chart using Aspose.Cells
// Description: This example creates an Excel workbook, populates task names with completion fractions, builds a stacked bar chart that acts as a progress bar, and configures each visible series to show only the percentage value (formatted to two decimal places) inside the bar. The workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | .NET | stacked bar chart | progress bar chart | data labels | percentage label | inside bar label | number format | Excel chart customization | chart series label
// Common Searches: Aspose.Cells show percentage on stacked bar chart | C# add data labels to progress bar chart Aspose | format chart data labels as 0.00% in .NET | place data labels inside bar Aspose.Cells example | display only percentage values in Excel chart using Aspose
// Developer Intent: Add data labels that display exact percentages to the visible series of a stacked‑bar progress chart.
// Use Cases: Project‑status dashboards that embed progress bars with precise percentages. | Automated Excel reports where each task’s completion is visualized with an inside‑bar label. | Weekly status sheets that require formatted percentage labels on stacked bar charts.
// AI Prompts: Generate C# code with Aspose.Cells to add percentage data labels (two decimal places) to a stacked bar chart and position them inside the bars. | Show how to hide value and category name labels while enabling only the percentage label for chart series in Aspose.Cells .NET. | Explain the steps to format chart data labels as 0.00% and set their position to InsideEnd using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsProgressBarDemo
{
    // This example creates an Excel workbook, populates task names with completion fractions, builds a stacked bar chart that acts as a progress bar, and configures each visible series to show only the percentage value (formatted to two decimal places) inside the bar. The workbook is saved as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data for a progress bar chart
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["A1"].PutValue("Task");
            dataSheet.Cells["B1"].PutValue("Completion");

            dataSheet.Cells["A2"].PutValue("Design");
            dataSheet.Cells["B2"].PutValue(0.30);   // 30%
            dataSheet.Cells["A3"].PutValue("Development");
            dataSheet.Cells["B3"].PutValue(0.55);   // 55%
            dataSheet.Cells["A4"].PutValue("Testing");
            dataSheet.Cells["B4"].PutValue(0.80);   // 80%
            dataSheet.Cells["A5"].PutValue("Deployment");
            dataSheet.Cells["B5"].PutValue(0.95);   // 95%

            // Add a chart sheet and create a stacked bar chart (used as a progress bar)
            Worksheet chartSheet = workbook.Worksheets[workbook.Worksheets.Add(SheetType.Chart)];
            int chartIndex = chartSheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 12);
            Chart chart = chartSheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("=Sheet1!$B$2:$B$5", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

            // Enable data labels for each visible series and show exact percentage values
            foreach (Series series in chart.NSeries)
            {
                // Show only the percentage value
                series.DataLabels.ShowPercentage = true;
                series.DataLabels.ShowValue = false;
                series.DataLabels.ShowCategoryName = false;

                // Use a number format that displays percentages with two decimal places
                series.DataLabels.NumberFormat = "0.00%";

                // Optional: position the label inside the bar for better readability
                series.DataLabels.Position = LabelPositionType.InsideEnd;
            }

            // Save the workbook
            workbook.Save("ProgressBarChartWithDataLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
