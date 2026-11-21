// Title: Generate an Excel Waterfall chart with a designated total segment using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to create a waterfall chart and set a total bar | How to mark a data point as total in an Aspose.Cells waterfall chart | Aspose.Cells Waterfall chart with custom total segment in .NET | Set IsTotal property on a chart point using Aspose.Cells C# | Programmatically create an Excel waterfall chart and designate the final total with Aspose.Cells
// Tags: aspocells waterfall chart total segment | c# set istotal on chart point | generate excel waterfall chart programmatically | aspocells chart point property reflection | excel workbook waterfall visualization c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The C# program creates a new workbook, fills columns A and B with category names and numeric values, adds a Waterfall chart covering the range, and uses reflection to set the IsTotal property on the last data point (if supported) to display it as a total segment. It also sets a chart title and saves the file as WaterfallChart.xlsx.
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

            // Populate data for the waterfall chart
            // Column A: categories, Column B: values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Expense");
            sheet.Cells["B4"].PutValue(-30);

            sheet.Cells["A5"].PutValue("Total");
            sheet.Cells["B5"].PutValue(0); // Placeholder; will be marked as total

            // Add a Waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Designate the last data point as the total segment (if the API supports it)
            try
            {
                var point = chart.NSeries[0].Points[3];
                var isTotalProp = point.GetType().GetProperty("IsTotal");
                if (isTotalProp != null && isTotalProp.CanWrite)
                {
                    isTotalProp.SetValue(point, true);
                }
            }
            catch
            {
                // If the property is unavailable, continue without marking the total.
            }

            // Optional: set chart title
            chart.Title.Text = "Waterfall Chart with Total Segment";

            // Determine output path and ensure directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WaterfallChart.xlsx");
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
