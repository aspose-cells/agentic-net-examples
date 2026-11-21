// Title: Programmatically set custom RGB colors for each series in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# how to change series color in a column chart | set individual series colors in Excel chart using Aspose.Cells .NET | apply custom RGB palette to chart series with Aspose.Cells API | C# code to assign red green blue colors to chart series in Aspose.Cells | customize chart series foreground color Aspose.Cells workbook
// Tags: Aspose.Cells chart series color customization | C# set series foreground color Aspose.Cells | custom RGB palette Excel chart Aspose.Cells | column chart series color assignment .NET | programmatic chart color styling Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, defines three RGB colors (red, green, blue), inserts a column chart, binds the data range, and assigns each custom color to the corresponding chart series by setting the series Area.ForegroundColor property before saving the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["D1"].PutValue("Series 3");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Define custom colors for the series
            Color series1Color = Color.FromArgb(255, 0, 0); // Red
            Color series2Color = Color.FromArgb(0, 255, 0); // Green
            Color series3Color = Color.FromArgb(0, 0, 255); // Blue

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Assign custom colors to each series
            chart.NSeries[0].Area.ForegroundColor = series1Color; // Series 1
            chart.NSeries[1].Area.ForegroundColor = series2Color; // Series 2
            chart.NSeries[2].Area.ForegroundColor = series3Color; // Series 3

            // Save the workbook with the chart
            string outputPath = "CustomPaletteChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
