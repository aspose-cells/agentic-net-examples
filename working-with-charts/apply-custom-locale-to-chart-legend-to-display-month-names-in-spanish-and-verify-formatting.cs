// Title: Generate a column chart with Spanish month names in the legend and verify blue, size‑12 legend formatting using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, adds a column chart with dates as categories, sets the chart locale to Spanish so the legend displays month names in Spanish, positions the legend at the bottom, and formats the legend font to blue size 12. | Include logic that checks whether the legend font color equals blue and the font size equals 12, then output the verification results.
// Common Searches: Aspose.Cells C# set chart legend locale to Spanish for month names | how to change legend font color and size in an Aspose.Cells column chart | verify chart legend formatting programmatically with Aspose.Cells .NET | display month names in Spanish on Excel chart legend using Aspose.Cells | apply custom culture to chart categories in Aspose.Cells C# example
// Tags: Aspose.Cells set chart legend locale Spanish | C# Aspose.Cells column chart legend font formatting | verify chart legend color and size Aspose.Cells | display month names in Spanish on Excel chart legend | apply custom culture to chart categories Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, inserts a column chart with date‑based categories, applies a Spanish locale so the legend shows month names in Spanish, positions the legend at the bottom, formats the legend font to blue size 12, verifies the formatting, and saves the file as ChartWithSpanishLegend.xlsx.
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

            // Populate sample data with dates (months) and values
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data series and category (month) data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Customize legend appearance for verification
            // Use LegendPositionType (compatible with older Aspose.Cells versions)
            chart.Legend.Position = LegendPositionType.Bottom;
            chart.Legend.Font.Color = Color.Blue;
            chart.Legend.Font.Size = 12;

            // Verify legend formatting
            bool isLegendBlue = chart.Legend.Font.Color.ToArgb() == Color.Blue.ToArgb();
            bool isLegendSize12 = chart.Legend.Font.Size == 12;
            Console.WriteLine($"Legend color is blue: {isLegendBlue}, size is 12: {isLegendSize12}");

            // Save the workbook
            string outputPath = "ChartWithSpanishLegend.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
