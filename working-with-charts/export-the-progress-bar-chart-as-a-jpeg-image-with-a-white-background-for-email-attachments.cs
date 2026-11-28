// Title: Export a stacked progress‑bar chart to a JPEG file with a white background using Aspose.Cells for .NET
// Common Searches: how to save an Aspose.Cells chart as a JPEG with white background in C# | export progress bar stacked bar chart to JPEG for email attachment using Aspose.Cells | set chart area background color before exporting to image with Aspose.Cells .NET | C# Aspose.Cells ImageOrPrintOptions example for JPEG quality and DPI | convert Excel chart to JPEG with white background using Aspose.Cells library
// Tags: Aspose.Cells chart to JPEG conversion | chart area white background Aspose.Cells | stacked bar progress chart export Aspose.Cells | ImageOrPrintOptions JPEG settings C# | export Excel chart as image for email Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // For ImageType enum

// The sample loads an existing workbook or creates one with a stacked bar chart that mimics a progress bar, sets the chart and plot areas to a white background, configures JPEG export options (quality 90, 150 DPI), and saves the first chart as a JPEG file suitable for email attachments.
class Program
{
    static void Main()
    {
        // Paths for the workbook and the output image.
        string workbookPath = "ProgressBarChart.xlsx";
        string outputImagePath = "ProgressBarChart.jpg";

        Workbook workbook;
        Worksheet sheet;

        // Load existing workbook if it exists; otherwise create a sample one.
        if (File.Exists(workbookPath))
        {
            try
            {
                workbook = new Workbook(workbookPath);
                sheet = workbook.Worksheets[0];
                Console.WriteLine($"Loaded workbook from \"{workbookPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }
        }
        else
        {
            try
            {
                workbook = new Workbook();
                sheet = workbook.Worksheets[0];

                // Sample data: Completed vs Remaining
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["B1"].PutValue("Completed");
                sheet.Cells["C1"].PutValue("Remaining");

                sheet.Cells["A2"].PutValue("Project");
                sheet.Cells["B2"].PutValue(70); // 70% completed
                sheet.Cells["C2"].PutValue(30); // 30% remaining

                // Add a stacked bar chart (acts as a progress bar)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 15, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Completed series (green)
                chart.NSeries.Add("B2:B2", true);
                chart.NSeries[0].Name = "Completed";
                chart.NSeries[0].Area.ForegroundColor = Color.Green;

                // Remaining series (light gray)
                chart.NSeries.Add("C2:C2", true);
                chart.NSeries[1].Name = "Remaining";
                chart.NSeries[1].Area.ForegroundColor = Color.LightGray;

                // Remove gaps to make it look like a single bar
                chart.GapWidth = 0;
                chart.GapDepth = 0;

                // Hide axes, legend and gridlines for a clean look
                chart.CategoryAxis.IsVisible = false;
                chart.ValueAxis.IsVisible = false;
                chart.ShowLegend = false;

                // Optional title
                chart.Title.Text = "Project Progress";

                Console.WriteLine("Created a sample progress‑bar chart because the source file was not found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating sample workbook: {ex.Message}");
                return;
            }
        }

        // Ensure there is at least one chart.
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        Chart targetChart = sheet.Charts[0];

        // Set white background for chart area and plot area.
        targetChart.ChartArea.Area.ForegroundColor = Color.White;
        targetChart.ChartArea.Border.IsVisible = false;
        targetChart.PlotArea.Area.ForegroundColor = Color.White;
        targetChart.PlotArea.Border.IsVisible = false;

        // Configure image export options.
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Jpeg,          // Use Aspose.Cells.Drawing.ImageType
            Quality = 90,                        // JPEG quality (0‑100)
            HorizontalResolution = 150,          // DPI
            VerticalResolution = 150
        };

        // Export the chart to a JPEG file.
        try
        {
            targetChart.ToImage(outputImagePath, imgOptions);
            Console.WriteLine($"Chart exported successfully to \"{outputImagePath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while exporting chart: {ex.Message}");
        }
    }
}
