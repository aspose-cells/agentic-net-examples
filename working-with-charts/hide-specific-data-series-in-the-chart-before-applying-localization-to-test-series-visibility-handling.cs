// Title: Hide a chart data series in Aspose.Cells C# with the IsFiltered property before localization
// Description: Creates a workbook, adds category labels and two series, builds a column chart, hides the first series by setting its IsFiltered flag, outputs the filtered series count, and saves the file as SeriesVisibilityDemo.xlsx.
// Keywords: Aspose.Cells hide chart series C# | IsFiltered chart series | filter data series Aspose.Cells | column chart series visibility .NET | chart series filtering Aspose
// Common Searches: How to hide a series in an Aspose.Cells chart using C# | IsFiltered property example Aspose.Cells | Count filtered series in Aspose.Cells chart | Hide first series before exporting workbook Aspose
// Developer Intent: Programmatically exclude a specific series from a chart to verify visibility handling prior to localization or export.
// Use Cases: Generate reports that show only selected series by filtering out others with IsFiltered. | Prevent hidden data from appearing in localized chart versions. | Validate the number of filtered series via chart.FilteredNSeries.Count before saving.
// AI Prompts: Write C# code that toggles the IsFiltered flag for multiple chart series based on runtime conditions using Aspose.Cells. | Explain how to refresh a chart after changing IsFiltered values in a .NET application. | Show how to retrieve and list the indexes or names of filtered series from an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    // Creates a workbook, adds category labels and two series, builds a column chart, hides the first series by setting its IsFiltered flag, outputs the filtered series count, and saves the file as SeriesVisibilityDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            // Category column
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            // First series values
            worksheet.Cells["B1"].PutValue("Series1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Second series values
            worksheet.Cells["C1"].PutValue("Series2");
            worksheet.Cells["C2"].PutValue(15);
            worksheet.Cells["C3"].PutValue(25);
            worksheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Add both series to the chart
            chart.NSeries.Add("B2:B4", true); // Series1
            chart.NSeries.Add("C2:C4", true); // Series2
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the first series using IsFiltered property
            chart.NSeries[0].IsFiltered = true;

            // Optional: display count of filtered series (should be 1)
            Console.WriteLine("Filtered series count: " + chart.FilteredNSeries.Count);

            // Save the workbook
            workbook.Save("SeriesVisibilityDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
