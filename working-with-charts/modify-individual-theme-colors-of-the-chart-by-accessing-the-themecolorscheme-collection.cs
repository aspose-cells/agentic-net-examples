// Title: Change individual chart theme colors in an Excel workbook using Aspose.Cells SetThemeColor (C#)
// Common Searches: c# change accent1 color of Excel chart using Aspose.Cells | how to customize chart series colors with Aspose.Cells theme palette | modify Excel workbook theme colors before adding a chart in .NET | Aspose.Cells example for column chart styling with custom theme colors | programmatically change Excel chart accent colors with C#
// Tags: Aspose.Cells SetThemeColor API | custom chart accent colors Excel C# | modify Excel theme palette programmatically | chart series theme colors Aspose.Cells | Excel workbook theme color customization .NET | SetThemeColor ThemeColorType usage | column chart theme color example Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ThemeColorChartDemo
{
    // The sample creates a workbook, sets custom theme colors (Accent1, Accent2, Text1, Background1) using SetThemeColor/GetThemeColor, adds a column chart that automatically adopts those colors, and saves the file as ThemeColorModifiedChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Populate sample data for the chart
                // -----------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");

                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // -----------------------------------------------------------------
                // 3. Modify individual theme colors that the chart will use.
                //    Aspose.Cells does not expose a direct Theme.ColorScheme collection,
                //    but the workbook provides SetThemeColor / GetThemeColor methods which
                //    affect the theme palette used by charts.
                // -----------------------------------------------------------------
                // Change Accent1 (used for the first series) to a custom teal color
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 150, 136));

                // Change Accent2 (used for the second series) to a custom amber color
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 193, 7));

                // Optionally change other theme entries (e.g., Text1, Background1)
                workbook.SetThemeColor(ThemeColorType.Text1, Color.DarkSlateGray);
                workbook.SetThemeColor(ThemeColorType.Background1, Color.WhiteSmoke);

                // Verify that the colors were set (output to console)
                Console.WriteLine("Accent1 set to: " + workbook.GetThemeColor(ThemeColorType.Accent1));
                Console.WriteLine("Accent2 set to: " + workbook.GetThemeColor(ThemeColorType.Accent2));

                // -----------------------------------------------------------------
                // 4. Add a column chart that will automatically pick the modified theme colors
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (both series)
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Give the chart a title
                chart.Title.Text = "Sales by Month (Theme Colors Modified)";

                // Optional: Force the chart to use the theme colors (default behavior)
                // No extra code needed; the series will use Accent1, Accent2, etc.

                // -----------------------------------------------------------------
                // 5. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "ThemeColorModifiedChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Any unexpected error will be printed to the console.
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
