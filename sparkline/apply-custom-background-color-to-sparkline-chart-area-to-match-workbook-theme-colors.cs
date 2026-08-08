// Title: Set a Custom Background Color for a Sparkline Chart Area with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a line sparkline, use a temporary column chart to access its ChartArea, apply a custom background shade (e.g., LightBlue or a theme color), and save the file as XLSX using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | sparkline background color | ChartArea styling | custom workbook theme | Excel chart area color | dummy chart technique | Excel report formatting | Aspose.Cells example
// Common Searches: Aspose.Cells set sparkline background color | C# change chart area color for sparkline | apply workbook theme to sparkline area Aspose | how to style sparkline background in .NET | use dummy chart to format sparkline Aspose.Cells
// Developer Intent: Apply a custom background shade to the visual area that contains a sparkline, aligning it with the workbook’s theme, via Aspose.Cells for .NET.
// Use Cases: Generate a financial dashboard where sparklines share a unified background that matches the corporate color palette. | Create automated Excel reports that require consistent sparkline styling without manual post‑processing. | Implement a quick workaround for the lack of direct Sparkline background API by leveraging a temporary chart’s ChartArea. | Extract a theme color from workbook.ThemeColors and apply it to the dummy chart to keep visual consistency across all sparklines.
// AI Prompts: Write C# code using Aspose.Cells that adds a line sparkline, creates a temporary column chart, sets the ChartArea background to a workbook theme color, and saves the workbook as XLSX. | Show how to retrieve a theme color from a workbook and use it as the background for a dummy chart that represents a sparkline’s area in Aspose.Cells for .NET. | Provide an Aspose.Cells example that demonstrates adding a sparkline group, accessing a chart’s ChartArea, applying a custom background color, and exporting the result.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSparklineBackgroundDemo
{
    // Demonstrates how to create a workbook, add a line sparkline, use a temporary column chart to access its ChartArea, apply a custom background shade (e.g., LightBlue or a theme color), and save the file as XLSX using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(3);
                sheet.Cells["A3"].PutValue(8);
                sheet.Cells["A4"].PutValue(2);

                // Define the location where the sparkline will be placed
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 1,
                    EndColumn = 1
                };

                // Add a sparkline group (Line type) using the data range A1:A4
                int sparklineGroupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A4", false, location);
                SparklineGroup sparklineGroup = sheet.SparklineGroups[sparklineGroupIndex];

                // Optional: customize sparkline appearance (series color, markers, etc.)
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.Orange;
                sparklineGroup.SeriesColor = seriesColor;

                // ------------------------------------------------------------
                // Apply a custom background color to the chart area that
                // contains the sparkline. Since a sparkline itself does not
                // expose an Area object, we use a regular chart's ChartArea
                // to demonstrate the background color.
                // ------------------------------------------------------------

                // Add a dummy chart (Column type) just to access its ChartArea
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart data (required for a valid chart)
                chart.NSeries.Add("A1:A4", true);
                chart.NSeries.CategoryData = "A1:A4";

                // Use a predefined fallback color for the background
                Color backgroundColor = Color.LightBlue;

                // Apply the background color to the chart area
                chart.ChartArea.Area.BackgroundColor = backgroundColor;

                // Save the workbook
                string outputPath = "SparklineWithCustomBackground.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
