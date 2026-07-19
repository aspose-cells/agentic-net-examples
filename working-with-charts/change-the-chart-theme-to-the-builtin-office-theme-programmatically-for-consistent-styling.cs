// Title: Apply the built‑in Office chart theme programmatically with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data and a column chart, then uses Workbook.SetThemeColor to apply the Office theme palette (background, text, accents, hyperlink) and sets an Office accent color on the chart series border. Saves as ChartWithOfficeTheme.xlsx.
// Keywords: Aspose.Cells | C# | Excel chart theme | Office theme | SetThemeColor | Workbook theme colors | Chart styling | ThemeColor API | programmatic chart formatting | Aspose.Cells example
// Common Searches: Aspose.Cells apply Office theme to chart | Set chart theme programmatically .NET | Workbook.SetThemeColor example C# | Change chart series border color Aspose.Cells | How to use Office theme palette in Aspose.Cells
// Developer Intent: Apply the built‑in Office theme to a chart for consistent corporate styling using Aspose.Cells.
// Use Cases: Generate Excel reports where all charts follow the standard Office theme for uniform appearance. | Customize chart series borders with Office accent colors to match corporate branding. | Programmatically enforce a predefined theme across multiple workbooks in an automated reporting pipeline.
// AI Prompts: Show me a C# snippet that uses Aspose.Cells to apply the Office theme to a workbook and its charts. | How can I set the Office Accent1 color on a chart series border with Aspose.Cells? | Explain the steps to switch between built‑in themes (Office, Metro) using Aspose.Cells in .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data and a column chart, then uses Workbook.SetThemeColor to apply the Office theme palette (background, text, accents, hyperlink) and sets an Office accent color on the chart series border. Saves as ChartWithOfficeTheme.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];
                chart.SetChartDataRange("A1:B4", false);

                // Apply the built‑in "Office" theme colors explicitly
                workbook.SetThemeColor(ThemeColorType.Background1, Color.FromArgb(255, 255, 255)); // White
                workbook.SetThemeColor(ThemeColorType.Text1, Color.FromArgb(0, 0, 0));           // Black
                workbook.SetThemeColor(ThemeColorType.Background2, Color.FromArgb(242, 242, 242)); // Light gray
                workbook.SetThemeColor(ThemeColorType.Text2, Color.FromArgb(89, 89, 89));        // Dark gray
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 112, 192));    // Blue
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 192, 0));    // Orange
                workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(112, 173, 71));   // Green
                workbook.SetThemeColor(ThemeColorType.Accent4, Color.FromArgb(255, 0, 0));      // Red
                workbook.SetThemeColor(ThemeColorType.Accent5, Color.FromArgb(255, 0, 255));    // Magenta
                workbook.SetThemeColor(ThemeColorType.Accent6, Color.FromArgb(0, 176, 80));     // Teal
                workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.FromArgb(0, 0, 255));    // Hyperlink blue
                workbook.SetThemeColor(ThemeColorType.FollowedHyperlink, Color.FromArgb(128, 0, 128)); // Purple

                // Demonstrate applying a theme color to the chart series
                if (chart.NSeries.Count > 0)
                {
                    chart.NSeries[0].Border.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
                    chart.NSeries[0].Border.Style = LineType.Solid;
                    chart.NSeries[0].Border.Weight = WeightType.MediumLine;
                }

                // Save the workbook
                string outputPath = "ChartWithOfficeTheme.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
