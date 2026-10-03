// Title: Create a weekly timeline line chart with bold axis titles and export it as a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Add a line chart to a worksheet, bind dates in column A and values in column B, set the category axis major unit to 7 days, make the chart title and axis titles bold, then save the workbook as a PDF. | Generate a timeline chart with weekly intervals, apply bold formatting to all titles, and output the result to a PDF file using Aspose.Cells in C#.
// Common Searches: aspnet create line chart with weekly dates and export to pdf using aspose.cells | how to set category axis major unit to 7 days in aspose.cells chart | bold chart title and axis titles in aspose.cells c# | save workbook as pdf after adding timeline chart asp.net | render weekly timeline chart in excel and convert to pdf with aspose.cells
// Tags: line chart weekly intervals Aspose.Cells | bold chart titles PDF export .NET | category axis major unit 7 days Aspose.Cells | timeline chart to PDF C# | date axis formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TimelinePdfExample
{
    // The example creates a new workbook, fills column A with weekly dates starting 1 Jan 2023 and column B with sample numeric values, adds a line chart that uses these ranges, sets the chart and axis titles to bold, configures the category axis to display dates at 7‑day intervals, and saves the workbook as Timeline.pdf.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with weekly dates and sample values.
                // Column A: Dates (weekly intervals)
                // Column B: Corresponding values
                DateTime startDate = new DateTime(2023, 1, 1);
                for (int i = 0; i < 10; i++)
                {
                    // Date (weekly interval)
                    sheet.Cells[i, 0].PutValue(startDate.AddDays(i * 7));
                    // Sample numeric value
                    sheet.Cells[i, 1].PutValue(10 + i * 5);
                }

                // Add a line chart (used as a timeline) to the worksheet.
                // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 30, 10);
                Chart timelineChart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                // Category (X) axis: dates in column A (A1:A10)
                // Values (Y) axis: numbers in column B (B1:B10)
                timelineChart.NSeries.Add("B1:B10", true);
                timelineChart.NSeries.CategoryData = "A1:A10";

                // Apply bold font to the chart title.
                timelineChart.Title.Text = "Weekly Timeline";
                timelineChart.Title.Font.IsBold = true;

                // Apply bold font to axis titles.
                timelineChart.CategoryAxis.Title.Text = "Week Starting";
                timelineChart.CategoryAxis.Title.Font.IsBold = true;

                timelineChart.ValueAxis.Title.Text = "Value";
                timelineChart.ValueAxis.Title.Font.IsBold = true;

                // Ensure the category axis displays dates with weekly intervals.
                // Set the major unit to 7 days.
                timelineChart.CategoryAxis.MajorUnit = 7;
                // Optionally format the date labels (if supported by the version).
                // timelineChart.CategoryAxis.NumberFormat = "yyyy-MM-dd";

                // Save the workbook as a PDF document.
                workbook.Save("Timeline.pdf", SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
