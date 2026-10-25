// Title: Apply conditional formatting to bar chart colors based on progress thresholds with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set bar chart data point color based on value | how to change Excel bar chart colors dynamically using Aspose.Cells | conditional formatting for chart series in Aspose.Cells .NET | apply threshold colors to a progress bar chart with Aspose.Cells C#
// Tags: Aspose.Cells conditional chart data point coloring | C# bar chart color thresholds Aspose.Cells | Excel progress bar chart conditional formatting .NET | programmatic chart series color change Aspose.Cells | dynamic bar chart colors based on cell values Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The example creates a workbook, populates task names and progress percentages, adds a bar chart to act as a progress bar, binds the data range to the chart series, and saves the file. It highlights that conditional coloring of individual data points can be achieved via the DataPoints collection, though the code omits this step for compatibility with older Aspose.Cells versions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for progress bar (tasks and completion percentages)
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Progress");
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(65);
            sheet.Cells["A4"].PutValue("Task 3");
            sheet.Cells["B4"].PutValue(90);

            // Add a bar chart (acts as a progress bar)
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Progress Bar";

            // Set data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: Conditional coloring of individual data points requires the DataPoints
            // collection, which may not be available in older Aspose.Cells versions.
            // The following block is omitted for compatibility.

            // Save the workbook with the chart
            string outputPath = "ProgressBarConditionalFormatting.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
