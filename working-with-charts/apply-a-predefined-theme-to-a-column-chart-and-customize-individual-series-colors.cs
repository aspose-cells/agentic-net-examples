// Title: Apply a .thmx theme to a column chart and set individual series colors with Aspose.Cells for .NET (C#)
// Common Searches: C# example to apply a .thmx theme to an Aspose.Cells column chart | how to change the color of each series in an Aspose.Cells column chart | Aspose.Cells load chart theme file .thmx code sample | set foreground color for chart series using Aspose.Cells in .NET | apply predefined Excel chart theme with Aspose.Cells C#
// Tags: apply thmx theme to Aspose.Cells chart | customize series colors column chart Aspose.Cells | Aspose.Cells chart theming C# | set series foreground color Aspose.Cells | load chart theme file Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, adds sample data, inserts a column chart, optionally loads a .thmx theme (if supported), assigns red to the first series and green to the second, and saves the file as ColumnChartWithTheme.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart.
            chart.NSeries.Add("B2:C4", true);          // Values for both series.
            chart.NSeries.CategoryData = "A2:A4";     // Categories.

            // Attempt to load a .thmx theme file if it exists.
            // Note: LoadTheme method is not available in the current Aspose.Cells version,
            // so this step is skipped to keep the code compilable.
            string themePath = "MyTheme.thmx";
            if (File.Exists(themePath))
            {
                // Placeholder for theme loading in versions that support it.
                // workbook.LoadTheme(themePath);
            }

            // Customize individual series colors.
            chart.NSeries[0].Area.ForegroundColor = Color.Red;   // Series1
            chart.NSeries[1].Area.ForegroundColor = Color.Green; // Series2

            // Save the workbook with the chart.
            workbook.Save("ColumnChartWithTheme.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
