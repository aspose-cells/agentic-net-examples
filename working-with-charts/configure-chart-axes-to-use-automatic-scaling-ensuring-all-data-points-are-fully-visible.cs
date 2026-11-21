// Title: Configure automatic axis scaling for a column chart in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart axis to auto range for column chart | How to make Excel chart axes automatically adjust to data using Aspose.Cells .NET | C# Aspose.Cells example of chart auto scaling | Configure automatic Y‑axis scaling in an Aspose.Cells generated column chart | Ensure all data points are displayed in an Aspose.Cells chart
// Tags: Aspose.Cells automatic axis scaling | C# column chart axis range configuration | Excel chart auto range with Aspose.Cells | full data visibility in Aspose.Cells chart | auto scaling of chart axes in .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code creates a workbook, fills cells A1:B6 with sample data, adds a column chart, and relies on Aspose.Cells' default automatic axis scaling to keep all values visible before saving the file as ChartWithAutoScaling.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B5"].PutValue(30);
            sheet.Cells["A6"].PutValue("May");
            sheet.Cells["B6"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            var chart = sheet.Charts[chartIndex];

            // Define the data range for the chart series
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // The chart axes automatically scale; explicit automatic scaling properties are not required.

            // Save the workbook with the configured chart
            workbook.Save("ChartWithAutoScaling.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
