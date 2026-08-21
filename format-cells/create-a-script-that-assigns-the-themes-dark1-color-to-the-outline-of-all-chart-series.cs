// Title: Set Dark1 Theme Color as Outline for All Chart Series with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, and uses Aspose.Cells to assign the Dark1 (Background1) theme color to each series border, makes the border visible, and saves the file.
// Keywords: Aspose.Cells chart series outline | Dark1 theme color | ThemeColorType.Background1 | C# chart border color | set series border Aspose.Cells | theme based chart styling .NET | chart series border visibility
// Common Searches: how to set chart series outline color using Aspose.Cells C# | apply Dark1 theme to all series borders in a chart | make chart series borders visible Aspose.Cells .NET | ThemeColorType.Background1 example for chart series | Aspose.Cells set series border to theme color
// Developer Intent: Apply the workbook’s Dark1 theme color to the outline of every chart series and ensure the borders are visible.
// Use Cases: Generate a column chart from worksheet data and style each series with the Dark1 theme for a unified look. | Create reporting workbooks that follow corporate branding by using a consistent theme‑based series outline. | Programmatically guarantee series borders are visible before distributing the XLSX file to end users.
// AI Prompts: Show C# code that loops through chart.NSeries in Aspose.Cells and sets series.Border.ThemeColor to ThemeColorType.Background1 with a zero tint, then makes the border visible. | Provide an example of applying the Dark1 theme color to the outline of all series in an Aspose.Cells chart, including required property settings. | Explain how to use ThemeColor with a tint/shade factor to style chart series borders consistently in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, and uses Aspose.Cells to assign the Dark1 (Background1) theme color to each series border, makes the border visible, and saves the file.
    public class AssignDark1ToSeriesOutline
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 15);
            }

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series
            chart.NSeries.Add("B2:C6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Assign the theme's Dark1 (Background1) color to the outline of each series
            foreach (Series series in chart.NSeries)
            {
                // ThemeColor constructor takes the theme color type and a tint/shade factor (0 = no change)
                series.Border.ThemeColor = new ThemeColor(ThemeColorType.Background1, 0);
                // Ensure the border is visible
                series.Border.IsVisible = true;
            }

            // Save the workbook
            workbook.Save("ChartSeriesDark1Outline.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                AssignDark1ToSeriesOutline.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
