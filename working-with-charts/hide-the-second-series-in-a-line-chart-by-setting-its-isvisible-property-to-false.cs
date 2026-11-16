// Title: Hide the second data series in a line chart with Aspose.Cells for .NET (C#)
// Common Searches: C# code to hide a specific series in an Aspose.Cells line chart | how to make a chart series invisible using Aspose.Cells .NET | remove second series from line chart Aspose.Cells example | Aspose.Cells hide chart series without deleting data | set chart series IsVisible false Aspose.Cells C#
// Tags: Aspose.Cells line chart series visibility C# | hide chart series Aspose.Cells .NET | remove series from NSeries Aspose.Cells | C# Aspose.Cells chart series manipulation | line chart data series control Aspose.Cells | Aspose.Cells chart series IsVisible property

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds sample data for two series, builds a line chart, and then hides the second series by removing it from the chart's NSeries collection, finally saving the file as an .xlsx workbook.
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

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (two series)
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the second series (index 1) by removing it from the collection
            if (chart.NSeries.Count > 1)
            {
                chart.NSeries.RemoveAt(1);
            }

            // Save the workbook with the chart
            string outputPath = "LineChart_HideSecondSeries.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
