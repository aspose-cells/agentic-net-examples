// Title: Hide a Series in a Stacked Bar Chart with Aspose.Cells for .NET
// Description: This example creates a workbook, adds category data and two series, builds a stacked bar chart, then hides the second series by setting chart.NSeries[1].IsFiltered to true. The workbook is saved as StackedBar_HideSeries.xlsx.
// Keywords: Aspose.Cells | C# | stacked bar chart | hide chart series | IsFiltered property | chart series visibility | filter chart series programmatically | Excel chart manipulation | remove series from chart | Aspose.Cells chart API
// Common Searches: Aspose.Cells hide series in chart | C# hide series stacked bar chart Aspose | IsFiltered chart series Aspose.Cells | filter out a series from Excel chart using code | programmatically hide chart series Aspose
// Developer Intent: Programmatically exclude a specific data series from a stacked bar chart so it does not appear in the generated Excel file.
// Use Cases: Simplify complex charts by hiding optional data series. | Create dynamic reports where series are shown/hidden based on user input. | Prepare chart templates with pre‑filtered series for later activation.
// AI Prompts: Show C# code to hide a series in an Aspose.Cells stacked bar chart using the IsFiltered property. | How can I toggle visibility of a chart series in Aspose.Cells for .NET? | Explain the difference between IsFiltered and IsVisible for chart series in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideSeriesInStackedBarChart
{
    // This example creates a workbook, adds category data and two series, builds a stacked bar chart, then hides the second series by setting chart.NSeries[1].IsFiltered to true. The workbook is saved as StackedBar_HideSeries.xlsx.
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

            // Series 2 values (this series will be hidden)
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

            // Hide the second series by marking it as filtered
            // When IsFiltered is true, the series will not be displayed on the chart
            chart.NSeries[1].IsFiltered = true;

            // Save the workbook
            workbook.Save("StackedBar_HideSeries.xlsx");
        }
    }
}
