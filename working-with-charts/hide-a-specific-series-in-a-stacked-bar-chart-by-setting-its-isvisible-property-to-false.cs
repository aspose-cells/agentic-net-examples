// Title: Hide a Series in a Stacked Bar Chart with Aspose.Cells for .NET
// Description: Creates a workbook, fills it with category and two data series, adds a stacked bar chart, and hides the second series by setting its IsFiltered property to false before saving the file.
// Keywords: Aspose.Cells | C# | .NET | stacked bar chart | hide chart series | IsFiltered property | chart series visibility | Excel chart manipulation | Aspose.Cells chart API | filter series without deleting data
// Common Searches: Aspose.Cells hide series stacked bar chart | C# hide chart series Aspose.Cells | IsFiltered property example Aspose.Cells | remove series from Excel chart programmatically | filter chart series Aspose.Cells .NET
// Developer Intent: Hide a specific series in a stacked bar chart so it does not appear in the generated workbook.
// Use Cases: Suppress a secondary data series to highlight the primary series in a stacked bar chart. | Toggle visibility of chart series based on user input without altering the underlying worksheet data.
// AI Prompts: Generate C# code using Aspose.Cells that hides the third series in a clustered column chart. | Show how to programmatically toggle visibility of multiple chart series at runtime with Aspose.Cells for .NET. | Provide an example of using the IsFiltered property to conditionally hide series across different chart types.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideSeriesInStackedBarChart
{
    // Creates a workbook, fills it with category and two data series, adds a stacked bar chart, and hides the second series by setting its IsFiltered property to false before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a stacked bar chart
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the second series (Series2) by marking it as filtered
            // When IsFiltered is true, the series will not be displayed on the chart
            chart.NSeries[1].IsFiltered = true;

            // Save the workbook
            workbook.Save("StackedBar_HideSeries.xlsx");
        }
    }
}
