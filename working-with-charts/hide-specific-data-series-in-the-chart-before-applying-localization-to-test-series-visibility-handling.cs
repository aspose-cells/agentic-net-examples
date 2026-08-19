// Title: Hide a chart series with IsFiltered in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, build a column chart, hide the second series by setting its IsFiltered property, verify the hidden series via the FilteredNSeries collection, and save the file—useful for testing chart visibility before localization.
// Keywords: Aspose.Cells C# chart series visibility | IsFiltered property Aspose.Cells | FilteredNSeries collection | hide chart series .NET | column chart series filter | chart localization testing | programmatic series hiding | Aspose.Cells chart filtering
// Common Searches: How to hide a chart series in Aspose.Cells using C# | Aspose.Cells IsFiltered example | Retrieve filtered series count with FilteredNSeries | Hide series before chart localization Aspose.Cells | Programmatically filter chart data in .NET
// Developer Intent: Exclude a specific data series from rendering in a chart so it does not appear during localization or export.
// Use Cases: Test how localization affects only visible chart data by temporarily hiding a series. | Programmatically remove low‑value or irrelevant series before generating reports. | Validate the number of hidden series using the FilteredNSeries collection for analytics.
// AI Prompts: Show code to hide multiple chart series in Aspose.Cells using IsFiltered. | Explain the FilteredNSeries collection and how to iterate over its items in C#. | Provide a sample that toggles series visibility based on user input in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    // Demonstrates how to create a workbook, add sample data, build a column chart, hide the second series by setting its IsFiltered property, verify the hidden series via the FilteredNSeries collection, and save the file—useful for testing chart visibility before localization.
    public class Program
    {
        public static void Main()
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add both series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the second series using the IsFiltered property
            chart.NSeries[1].IsFiltered = true;

            // Verify the filtered series collection
            SeriesCollection filtered = chart.FilteredNSeries;
            Console.WriteLine("Filtered series count after hiding Series2: " + filtered.Count);

            // Save the workbook
            workbook.Save("SeriesVisibilityDemo.xlsx");
        }
    }
}
