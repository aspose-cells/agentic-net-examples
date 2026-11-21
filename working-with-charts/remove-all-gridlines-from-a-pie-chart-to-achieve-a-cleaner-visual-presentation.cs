// Title: Hide all gridlines on a pie chart with Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove gridlines from a pie chart using Aspose.Cells | Aspose.Cells hide major and minor gridlines on chart axes | How to create a pie chart without axis gridlines in Excel via Aspose.Cells | Disable chart gridlines in a generated .xlsx file with Aspose.Cells C# | Remove value and category axis gridlines from Excel pie chart programmatically
// Tags: hide chart gridlines Aspose.Cells C# | pie chart axis gridlines removal .NET | Aspose.Cells chart formatting without gridlines | disable major minor gridlines Excel chart API | generate pie chart without gridlines Aspose.Cells | Excel workbook save without chart gridlines C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a pie chart, and disables both major and minor gridlines on the value and category axes before saving the file as PieChart_NoGridlines.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide gridlines on both value and category axes for a cleaner look
            chart.ValueAxis.MajorGridLines.IsVisible = false;
            chart.ValueAxis.MinorGridLines.IsVisible = false;
            chart.CategoryAxis.MajorGridLines.IsVisible = false;
            chart.CategoryAxis.MinorGridLines.IsVisible = false;

            // Prepare output path
            string outputPath = "PieChart_NoGridlines.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
