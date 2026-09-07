// Title: Thicken progress bar chart bars by lowering GapWidth with Aspose.Cells for .NET
// AI Prompts: Generate an Excel workbook, add a bar chart, and set its GapWidth property to 20% to create thicker progress bars using Aspose.Cells. | Modify an existing Aspose.Cells chart to decrease the GapWidth value, making the bars appear wider in a progress‑bar visualization. | Write a C# program that builds a progress‑bar style bar chart and customizes bar thickness through the chart.GapWidth setting.
// Common Searches: Aspose.Cells how to reduce chart gap width for thicker bars in C# | C# code example to make progress bar chart bars wider using Aspose.Cells | adjust bar spacing in Excel bar chart with Aspose.Cells .NET | set GapWidth percentage for bar chart in Aspose.Cells tutorial | increase visual weight of progress bar chart columns Aspose.Cells
// Tags: Aspose.Cells chart GapWidth adjustment | thicker bars in Excel bar chart .NET | progress bar chart styling Aspose.Cells | customize bar spacing Excel chart C# | set chart bar thickness Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program creates a workbook, inserts task data, adds a bar chart used as a progress bar, sets chart.GapWidth to 20% for noticeably thicker bars, and saves the file as ProgressBarChart.xlsx.
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

            // Populate sample data for the progress bar chart
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Progress");
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["B2"].PutValue(30);   // 30%
            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["B3"].PutValue(70);   // 70%

            // Add a bar chart (used as a progress bar)
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Adjust the gap width to make the bars appear thicker
            // GapWidth is a percentage (default 150). Lower values produce thicker bars.
            chart.GapWidth = 20; // 20% gap width makes bars noticeably thicker

            // Define output file path
            string outputPath = "ProgressBarChart.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
