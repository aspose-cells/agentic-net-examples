// Title: Programmatically Apply the Built‑in Office Chart Theme with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data and a column chart, sets the Office palette via Workbook.SetThemeColor, assigns Accent1 to the series border, applies a simple Office‑style layout with Chart.Style, and saves the file as ChartWithOfficeTheme.xlsx.
// Keywords: Aspose.Cells C# chart theme | SetThemeColor Office palette | apply Office theme Aspose.Cells | chart.Style built‑in theme | programmatic Excel chart styling .NET | ThemeColorType Accent1 | column chart Aspose.Cells | Excel workbook theme colors
// Common Searches: how to set Office theme on a chart using Aspose.Cells C# | Aspose.Cells SetThemeColor example | apply built‑in chart style programmatically .NET | change chart colors to Office palette Aspose.Cells | C# code for chart theme with Aspose.Cells
// Developer Intent: Apply the built‑in Office theme to a chart in an Aspose.Cells workbook using C#.
// Use Cases: Standardize corporate Excel reports so every chart follows the Office color scheme. | Automate bulk updates of existing workbooks to enforce a consistent chart appearance. | Generate dashboards with predefined Office‑style charts for scheduled data exports.
// AI Prompts: Show C# code that uses Aspose.Cells to set Office theme colors on a workbook and apply them to a column chart. | Explain how Workbook.SetThemeColor and Chart.Style work together to create an Office‑styled chart in Aspose.Cells. | Provide step‑by‑step instructions for programmatically applying the built‑in Office theme to all charts in a .NET Excel file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Creates a workbook, adds sample data and a column chart, sets the Office palette via Workbook.SetThemeColor, assigns Accent1 to the series border, applies a simple Office‑style layout with Chart.Style, and saves the file as ChartWithOfficeTheme.xlsx.
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
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", false);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply the built‑in “Office” theme colors to the workbook
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(68, 114, 196));
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(237, 125, 49));
                workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(165, 165, 165));
                workbook.SetThemeColor(ThemeColorType.Accent4, Color.FromArgb(255, 192, 0));
                workbook.SetThemeColor(ThemeColorType.Accent5, Color.FromArgb(91, 155, 213));
                workbook.SetThemeColor(ThemeColorType.Accent6, Color.FromArgb(112, 173, 71));
                workbook.SetThemeColor(ThemeColorType.Text1, Color.Black);
                workbook.SetThemeColor(ThemeColorType.Background1, Color.White);
                workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.FromArgb(5, 99, 193));
                workbook.SetThemeColor(ThemeColorType.FollowedHyperlink, Color.FromArgb(149, 79, 114));

                // Apply theme colors to the first series of the chart
                Series series = chart.NSeries[0];
                series.Border.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
                series.Area.FillFormat.FillType = FillType.Solid;
                // Note: SolidFillColor property may not be available in older versions; omitted for compatibility.

                // Optionally set a built‑in chart style that aligns with the Office theme
                chart.Style = 1; // Simple Office‑style layout

                // Save the workbook
                string outputPath = "ChartWithOfficeTheme.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
