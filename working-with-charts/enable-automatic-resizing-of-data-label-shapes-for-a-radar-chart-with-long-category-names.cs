// Title: Automatically resize radar chart data label shapes for long category names using Aspose.Cells for .NET
// AI Prompts: Set series.DataLabels.IsAutoFit = true so radar chart data label shapes expand automatically to fit long category text in Aspose.Cells C#. | Add code that enables auto‑fit for data labels on a radar chart before saving the workbook with Aspose.Cells. | Configure the radar chart's data label auto‑size property in a C# Aspose.Cells example to handle lengthy category names.
// Common Searches: Aspose.Cells C# enable auto fit for radar chart data labels with long category names | how to make radar chart data labels adjust size automatically in .NET | resize data label shapes for long category text in Aspose.Cells radar chart | C# Aspose.Cells set data label auto size on radar chart
// Tags: auto fit chart data labels Aspose.Cells | radar chart label auto size C# | long category names radar chart Aspose.Cells | set IsAutoFit property chart series

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a workbook, fills column A with very long category names and column B with values, adds a radar chart, binds the series and categories, and shows values on data labels. To ensure the label shapes automatically resize for the long category names, set the series' DataLabels.IsAutoFit property to true before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Fill data with long category names
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["A2"].PutValue("VeryLongCategoryNameOne");
            sheet.Cells["A3"].PutValue("ExtremelyLongCategoryNameTwo");
            sheet.Cells["A4"].PutValue("SuperLongCategoryNameThree");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a radar chart
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Show values on data labels
            var series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Save the workbook
            workbook.Save("RadarChartAutoResize.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
