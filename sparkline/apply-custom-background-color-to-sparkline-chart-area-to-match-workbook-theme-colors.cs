// Title: C# Example: Apply Workbook Accent1 Theme Color to Sparkline and Chart Background with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a line sparkline group, retrieve the workbook's Accent1 theme color, set the sparkline series color, and apply the same color to a chart area's background before saving the file.
// Keywords: Aspose.Cells sparkline background color | C# sparkline theme color | GetThemeColor Aspose.Cells | Apply Accent1 color to chart area | Aspose.Cells example sparkline | Workbook theme color C# | Aspose.Cells chart styling
// Common Searches: how to set sparkline color using workbook theme in Aspose.Cells | C# Aspose.Cells apply Accent1 theme to chart background | retrieve theme color for sparkline Aspose.Cells | set sparkline series color from workbook theme | Aspose.Cells example for themed chart area
// Developer Intent: Use the workbook's Accent1 theme color for both sparkline series and chart area background.
// Use Cases: Create a sparkline group that automatically matches the workbook's theme. | Synchronize sparkline and chart colors for consistent visual design. | Generate reports where theme colors adapt to different corporate branding.
// AI Prompts: Show C# code to get the Accent1 theme color from a workbook and apply it to a sparkline using Aspose.Cells. | Provide an Aspose.Cells example that sets both a sparkline series color and a chart area background to the same theme color. | Explain how GetThemeColor and CreateCellsColor work together to theme sparkline and chart visuals in .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a line sparkline group, retrieve the workbook's Accent1 theme color, set the sparkline series color, and apply the same color to a chart area's background before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            worksheet.Cells["A1"].PutValue(5);
            worksheet.Cells["A2"].PutValue(3);
            worksheet.Cells["A3"].PutValue(8);
            worksheet.Cells["A4"].PutValue(2);

            // Define the location range for the sparkline (cells B1:B4)
            CellArea location = new CellArea
            {
                StartRow = 0,   // Row 1 (zero‑based)
                EndRow = 3,     // Row 4
                StartColumn = 1, // Column B (zero‑based)
                EndColumn = 1   // Column B
            };

            // Add a sparkline group of type Line
            int sparklineGroupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:A4", false, location);
            SparklineGroup sparklineGroup = worksheet.SparklineGroups[sparklineGroupIndex];

            // Retrieve a theme color (Accent1) from the workbook's theme
            Color themeAccentColor = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Apply the theme color to the sparkline series
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = themeAccentColor;
            sparklineGroup.SeriesColor = seriesColor;

            // OPTIONAL: Apply the same theme color to a regular chart's background
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("A1:A4", true);
            chart.ChartArea.Area.BackgroundColor = themeAccentColor;

            // Save the workbook
            workbook.Save("SparklineWithThemeBackground.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
