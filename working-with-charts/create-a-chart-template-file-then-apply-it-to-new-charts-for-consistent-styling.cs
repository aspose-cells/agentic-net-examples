// Title: Create and Apply a Reusable Chart Template with Aspose.Cells for .NET
// Description: Demonstrates how to generate a workbook, add sample data, build a column chart, apply a built‑in style, set a custom title, legend, and foreground colors for the plot and chart areas, switch to a 3‑D clustered column while preserving formatting, and save the file as XLSX. The same styling can be saved as a template and reused for other charts to ensure visual consistency.
// Keywords: Aspose.Cells chart template | C# chart styling | apply built‑in chart style .NET | change chart type preserve formatting | chart area foreground color | plot area color Aspose.Cells | save workbook with chart | reusable Excel chart format | Aspose.Cells example
// Common Searches: how to create a chart template with Aspose.Cells | apply same style to multiple charts .NET | change chart type without losing formatting Aspose.Cells | set chart area background color using Aspose.Cells | save Excel file with styled chart C#
// Developer Intent: Generate a styled chart once and reuse its formatting across additional charts, then export the workbook.
// Use Cases: Produce a column chart from a data range and apply a predefined visual theme. | Define custom foreground colors for chart and plot areas to match corporate branding. | Switch a chart to a different type (e.g., 3‑D clustered column) while retaining all prior styling. | Save the styled workbook as an XLSX file for distribution or further processing.
// AI Prompts: Show me how to save a chart's formatting as a template file with Aspose.Cells and reuse it for new charts. | Give an example of preserving all style settings when converting a column chart to a 3‑D chart in Aspose.Cells for C#. | Explain how to customize chart area and plot area colors programmatically using Aspose.Cells.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplateDemo
{
    // Demonstrates how to generate a workbook, add sample data, build a column chart, apply a built‑in style, set a custom title, legend, and foreground colors for the plot and chart areas, switch to a 3‑D clustered column while preserving formatting, and save the file as XLSX. The same styling can be saved as a template and reused for other charts to ensure visual consistency.
    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // Step 1: Create a workbook and add a chart with desired styling
                // -----------------------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Sample data for the chart
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["A2"].PutValue("X");
                ws.Cells["A3"].PutValue("Y");
                ws.Cells["A4"].PutValue("Z");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["B2"].PutValue(15);
                ws.Cells["B3"].PutValue(25);
                ws.Cells["B4"].PutValue(35);

                // Add a chart
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = ws.Charts[chartIdx];

                // Set data range for the chart
                chart.SetChartDataRange("A1:B4", false);

                // Apply styling (similar to a template)
                chart.Style = 7; // Example builtin style
                chart.Title.Text = "Styled Chart";
                chart.ShowLegend = true;
                chart.PlotArea.Area.ForegroundColor = Color.LightYellow;
                chart.ChartArea.Area.ForegroundColor = Color.LightBlue;

                // Optionally change the chart type while retaining the styling
                chart.Type = ChartType.Column3DClustered;

                // -----------------------------------------------------------------
                // Step 2: Save the workbook containing the styled chart
                // -----------------------------------------------------------------
                string outputPath = "WorkbookWithStyledChart.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                wb.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine("Chart created and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
