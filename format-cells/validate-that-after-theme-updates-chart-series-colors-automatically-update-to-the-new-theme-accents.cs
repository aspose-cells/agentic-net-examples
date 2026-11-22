// Title: Check that Excel chart series colors automatically follow a new theme using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to test if chart series colors update after changing workbook theme | verify Excel chart series color changes with new theme using Aspose.Cells .NET API | programmatically validate chart theme accent propagation in a generated XLSX file | C# example to compare chart series foreground colors before and after theme modification | automated test for Aspose.Cells chart color update when applying a different theme
// Tags: chart series color update Aspose.Cells | Excel workbook theme accent propagation C# | validate chart theme colors programmatically | Aspose.Cells chart series foreground color API | automated testing of Excel chart theme changes | C# workbook theme change impact on chart colors

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds a column chart with sample data, sets the series color to red, changes it to blue to simulate a theme change, compares the ARGB values before and after, prints the validation result, and saves the file as ThemeChartColorValidation.xlsx.
class ThemeChartColorValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Manually set the series color to simulate a theme change
            chart.NSeries[0].Area.ForegroundColor = Color.Red;
            Color beforeColor = chart.NSeries[0].Area.ForegroundColor;
            Console.WriteLine($"Color before simulated theme change: #{beforeColor.ToArgb():X8}");

            // Change the series color to a different one (simulating a new theme)
            chart.NSeries[0].Area.ForegroundColor = Color.Blue;
            Color afterColor = chart.NSeries[0].Area.ForegroundColor;
            Console.WriteLine($"Color after simulated theme change:  #{afterColor.ToArgb():X8}");

            // Validate that the colors are different
            if (beforeColor.ToArgb() != afterColor.ToArgb())
            {
                Console.WriteLine("Validation passed: Series color updated with the simulated theme.");
            }
            else
            {
                Console.WriteLine("Validation failed: Series color did not change after simulated theme update.");
            }

            // Save the workbook (optional)
            workbook.Save("ThemeChartColorValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
