// Title: Enable text wrapping for pie‑chart data labels with Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net how to wrap pie chart data label text using Aspose.Cells | set DataLabels.IsTextWrapped property on a pie chart series in Aspose.Cells | example of enabling text wrap for Excel chart data labels with Aspose.Cells C# | wrap long category names in pie chart labels using Aspose.Cells .NET | Aspose.Cells chart label formatting to prevent truncation in pie charts
// Tags: Aspose.Cells pie chart label text wrap C# | DataLabels.IsTextWrapped property example .NET | Excel chart data label formatting Aspose.Cells | prevent truncation of pie chart labels C# | chart series label wrap Aspose.Cells | chart label style customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds category and value data, inserts a pie chart, defines the series range, enables data labels to show values, sets DataLabels.IsTextWrapped to true so label text wraps automatically, and saves the file as PieChartWithWrappedLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(50);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Add series with both category and value ranges (comma‑separated)
            chart.NSeries.Add("A2:A4,B2:B4", false);

            // Enable data labels and wrap text for the series
            var series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // display values
            series.DataLabels.IsTextWrapped = true;      // wrap long labels

            // Save the workbook
            workbook.Save("PieChartWithWrappedLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
