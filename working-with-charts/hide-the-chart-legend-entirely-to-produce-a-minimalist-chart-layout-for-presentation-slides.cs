// Title: Hide the legend of a column chart in Excel using Aspose.Cells for .NET (C#) to create a clean slide‑ready chart
// Common Searches: C# Aspose.Cells hide chart legend for column chart | remove legend from Excel chart using Aspose.Cells .NET | how to delete chart legend in Aspose.Cells C# example | create minimalist Excel chart without legend using Aspose.Cells | Aspose.Cells chart.Legend.IsDeleted property usage in C#
// Tags: Aspose.Cells chart legend deletion | C# Excel column chart without legend | Aspose.Cells hide legend property | minimalist Excel chart Aspose.Cells | Aspose.Cells chart formatting C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds sample data, inserts a column chart, sets the data range, hides the chart legend by setting chart.Legend.IsDeleted to true, and saves the workbook as an XLSX file, producing a minimalist chart suitable for presentation slides.
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend completely
            chart.Legend.IsDeleted = true;

            // Verify that the legend is hidden
            Console.WriteLine("Legend hidden: " + chart.Legend.IsDeleted);

            // Save the workbook to a file
            string outputPath = "ChartWithoutLegend.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
