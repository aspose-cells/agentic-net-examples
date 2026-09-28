// Title: Create a horizontal progress bar chart in Excel using Aspose.Cells for .NET by stacking a visible progress series over a hidden total series
// AI Prompts: Generate a stacked horizontal bar chart where the total series is invisible and the progress series is colored green with Aspose.Cells in C#. | Set the background series fill to transparent and hide its border, then assign task names as category labels for the chart. | Remove the legend, add a chart title, and save the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: how to make a progress bar chart in Excel using Aspose.Cells C# stacked bar | hide total series in Aspose.Cells stacked bar chart for progress visualization | horizontal progress bar with transparent background series Aspose.Cells .NET example
// Tags: Aspose.Cells stacked bar chart progress visualization | transparent background series Aspose.Cells | green fill progress series Excel .NET | horizontal progress bar chart C# | remove legend Aspose.Cells chart

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates an Excel workbook, populates task data, and adds a horizontal stacked bar chart. The total series is made transparent to act as a hidden background, while the progress series is displayed in green. Task names are used as category labels, the legend is hidden, a title is set, and the file is saved as ProgressBarChart.xlsx.
class ProgressBarChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row.
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Total");
            sheet.Cells["C1"].PutValue("Progress");

            // Sample tasks.
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["A4"].PutValue("Task 3");
            sheet.Cells["A5"].PutValue("Task 4");

            // Totals (100% for each task).
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(100);
            sheet.Cells["B4"].PutValue(100);
            sheet.Cells["B5"].PutValue(100);

            // Completed percentages.
            sheet.Cells["C2"].PutValue(70);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(90);
            sheet.Cells["C5"].PutValue(30);

            // Add a stacked bar chart (horizontal) positioned at D2, spanning 15 rows x 10 columns.
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 1, 3, 15, 13);
            Chart chart = sheet.Charts[chartIndex];

            // Background series (total) – will be hidden.
            chart.NSeries.Add("B2:B5", true);
            // Visible progress series.
            chart.NSeries.Add("C2:C5", true);

            // Hide the background series by making it transparent.
            Series backgroundSeries = chart.NSeries[0];
            backgroundSeries.Area.ForegroundColor = Color.Transparent;
            backgroundSeries.Border.IsVisible = false;

            // Format the progress series with a solid fill color.
            Series progressSeries = chart.NSeries[1];
            progressSeries.Area.ForegroundColor = Color.Green;
            progressSeries.Border.IsVisible = false;

            // Set category axis labels to task names.
            chart.NSeries.CategoryData = "A2:A5";

            // Remove the legend (optional).
            chart.ShowLegend = false;

            // Set chart title.
            chart.Title.Text = "Progress Bar Chart";

            // Save the workbook.
            workbook.Save("ProgressBarChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
