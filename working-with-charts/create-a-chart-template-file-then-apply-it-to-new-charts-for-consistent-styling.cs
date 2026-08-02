// Title: Create and Style a Column Chart, Add Gradient Fill, Convert to 3D, and Save with Aspose.Cells for .NET
// Description: This example shows how to generate a workbook, populate cells A1:B4, add a column chart, apply built‑in style 5, set a radial gradient fill on the first series, switch the chart to a 3D clustered column while preserving formatting, and finally save the file as WorkbookWithStyledChart.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart template .NET | C# create column chart Aspose.Cells | gradient fill chart series Aspose.Cells | convert chart to 3D Aspose.Cells | save workbook with styled chart | apply built‑in chart style | chart styling automation
// Common Searches: How to add gradient fill to a chart series in Aspose.Cells C# | Change chart type without losing formatting Aspose.Cells | Save a workbook that contains a styled chart using Aspose.Cells | Create a reusable chart template with Aspose.Cells | Apply a chart template to multiple workbooks in .NET
// Developer Intent: Programmatically build a workbook, insert a column chart with custom styling, transform it to a 3D chart, and persist the result.
// Use Cases: Standardize visual branding across quarterly sales reports by reusing the same chart style and gradient fill. | Dynamically switch a 2D column chart to a 3D view in a dashboard without re‑applying formatting. | Generate a library of chart templates that can be applied to new workbooks for consistent appearance.
// AI Prompts: Show C# code to export the styled chart as a chart template file (.cst) with Aspose.Cells. | Demonstrate how to load a saved chart template and apply it to a new chart in a different workbook. | Explain how to copy gradient fill settings from one chart series to another using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTemplateDemo
{
    // This example shows how to generate a workbook, populate cells A1:B4, add a column chart, apply built‑in style 5, set a radial gradient fill on the first series, switch the chart to a 3D clustered column while preserving formatting, and finally save the file as WorkbookWithStyledChart.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and add a chart with styling
                // -------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Sample data for the chart
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["A3"].PutValue("B");
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["B2"].PutValue(15);
                ws.Cells["B3"].PutValue(25);
                ws.Cells["B4"].PutValue(35);

                try
                {
                    // Add a column chart
                    int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart chart = ws.Charts[chartIdx];

                    // Set data range
                    chart.SetChartDataRange("A1:B4", false);

                    // Apply a built‑in style
                    chart.Style = 5;

                    // Apply a gradient fill to the first series
                    Series series = chart.NSeries[0];
                    series.Area.FillFormat.FillType = FillType.Gradient;
                    series.Area.FillFormat.GradientFill.SetPresetThemeGradient(
                        PresetThemeGradientType.RadialGradient,
                        ThemeColorType.Accent1);

                    // Change the chart type while keeping the styling
                    chart.Type = ChartType.Column3DClustered;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating or styling chart: {ex.Message}");
                    return;
                }

                // -------------------------------------------------
                // 2. Save the workbook
                // -------------------------------------------------
                string outputPath = "WorkbookWithStyledChart.xlsx";
                try
                {
                    wb.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
