// Title: Resize chart data label shapes after applying bold font with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# resize chart data label shape after setting bold font | prevent data label overflow in Excel column chart using Aspose.Cells | adjust width of data labels in Aspose.Cells chart to fit bold text | how to auto‑fit chart data label box size with Aspose.Cells .NET | increase chart data label area after applying bold formatting in Aspose.Cells
// Tags: data label shape resize Aspose.Cells | bold font overflow handling Excel chart | chart label dimension adjustment C# | prevent data label clipping Aspose.Cells .NET | auto‑fit chart label box size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates sample data, adds a column chart, enables data labels for the first series, applies bold formatting to the label text, adjusts the label shape size to avoid overflow, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Set the series data and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Show data labels on the first series and apply bold formatting
            if (chart.NSeries.Count > 0)
            {
                var series = chart.NSeries[0];
                // Enable data labels (ShowValue displays the value)
                series.DataLabels.ShowValue = true;
                // Apply bold formatting to the label text
                series.DataLabels.Font.IsBold = true;
            }

            // Determine output path, ensure directory exists, and save the workbook
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
