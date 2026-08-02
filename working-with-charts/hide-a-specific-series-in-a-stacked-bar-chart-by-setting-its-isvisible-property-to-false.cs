// Title: Hide a Series in a Stacked Bar Chart with Aspose.Cells for .NET (IsFiltered)
// Description: Creates a workbook, adds category data and two series, builds a stacked bar chart, then hides the second series by setting its IsFiltered flag to true before saving the file.
// Keywords: Aspose.Cells | C# chart series hide | IsFiltered | stacked bar chart | .NET Excel chart | chart series visibility | hide series Aspose.Cells | Excel chart filtering | Aspose.Cells API
// Common Searches: Aspose.Cells hide chart series C# | How to hide a series in a stacked bar chart using Aspose.Cells | Set IsFiltered true Aspose.Cells chart | Remove series from Excel chart programmatically | Chart series visibility .NET Aspose
// Developer Intent: Programmatically exclude a specific data series from a stacked bar chart so it does not render in the exported Excel workbook.
// Use Cases: Generate reports where optional series are hidden based on user selections. | Create chart templates with placeholder series that can be suppressed until real data is supplied. | Export Excel files that contain confidential data series, keeping the data but preventing visual display.
// AI Prompts: Write C# code with Aspose.Cells that hides a chosen series in a stacked bar chart by setting IsFiltered to true. | Show how to toggle chart series visibility in Aspose.Cells based on a boolean parameter. | Explain the difference between IsFiltered and IsVisible for Aspose.Cells chart series and when to use each.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideSeriesInStackedBarChart
{
    // Creates a workbook, adds category data and two series, builds a stacked bar chart, then hides the second series by setting its IsFiltered flag to true before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // Series 1 values
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(40);
            sheet.Cells["B4"].PutValue(50);

            // Series 2 values (the series we will hide)
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a stacked bar chart
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the second series (Series2) by marking it as filtered
            // When IsFiltered is true, the series will not be displayed on the chart
            chart.NSeries[1].IsFiltered = true;

            // Save the workbook
            workbook.Save("HiddenSeriesStackedBarChart.xlsx");
        }
    }
}
