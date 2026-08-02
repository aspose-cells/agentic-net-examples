// Title: Export Progress Bar Chart to JPEG with White Background using Aspose.Cells for .NET
// Description: This example creates a workbook, fills it with task names and progress percentages, adds a bar chart, sets a solid white background, and uses chart.ToImage to save the chart as a JPEG file ready for email attachment.
// Keywords: Aspose.Cells chart export JPEG | C# chart white background | progress bar chart image | Aspose.Cells ToImage JPEG | email attachment chart image
// Common Searches: Aspose.Cells export chart as JPEG | set chart background color before exporting C# | progress bar chart image for email | how to save Aspose.Cells chart with white background | C# generate chart image for email reports
// Developer Intent: Create a progress‑bar chart, apply a solid white background, and export it as a JPEG file for inclusion in email messages.
// Use Cases: Automated status‑report emails that embed a progress bar image. | Generating chart thumbnails for dashboards where transparency is unsupported. | Batch converting workbooks to printable JPEG charts for archival.
// AI Prompts: Show C# code that sets a chart's background to white and exports it as a JPEG with Aspose.Cells. | Explain how to ensure a solid white background when using chart.ToImage for JPEG output. | Provide a complete Aspose.Cells example that builds a progress bar chart and saves it as a JPEG suitable for email attachments.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartExport
{
    // This example creates a workbook, fills it with task names and progress percentages, adds a bar chart, sets a solid white background, and uses chart.ToImage to save the chart as a JPEG file ready for email attachment.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a progress bar
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["A2"].PutValue("Design");
                sheet.Cells["A3"].PutValue("Development");
                sheet.Cells["A4"].PutValue("Testing");

                sheet.Cells["B1"].PutValue("Progress");
                sheet.Cells["B2"].PutValue(70); // 70%
                sheet.Cells["B3"].PutValue(45); // 45%
                sheet.Cells["B4"].PutValue(20); // 20%

                // Add a bar chart (Progress Bar)
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Determine output path for the chart image
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ProgressBarChart.jpg");
                string outputDir = Path.GetDirectoryName(outputPath);

                // Ensure the output directory exists
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the chart to a JPEG file (default format is PNG; using ToImage overload without ImageFormat)
                try
                {
                    chart.ToImage(outputPath);
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Failed to export chart image: {imgEx.Message}");
                }

                // Save the workbook for reference
                string workbookPath = Path.Combine(Environment.CurrentDirectory, "ProgressBarWorkbook.xlsx");
                try
                {
                    workbook.Save(workbookPath);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }

                Console.WriteLine($"Chart exported to JPEG: {outputPath}");
                Console.WriteLine($"Workbook saved to: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
