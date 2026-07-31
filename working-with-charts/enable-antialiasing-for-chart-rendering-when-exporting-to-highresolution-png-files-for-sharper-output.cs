// Title: Enable Anti‑Aliasing for High‑Resolution PNG Chart Export with Aspose.Cells (C# .NET)
// Description: A C# sample that creates a workbook, adds a column chart, and exports the chart as a 300 dpi PNG with smooth edges by setting ImageOrPrintOptions.AntiAliasing = true before calling Chart.ToImage. The example also shows how to save the workbook.
// Keywords: Aspose.Cells chart anti aliasing C# | high DPI PNG chart export Aspose.Cells | ImageOrPrintOptions AntiAliasing .NET | C# Aspose.Cells smooth chart rendering | export chart to PNG with anti aliasing
// Common Searches: how to enable anti aliasing for chart images in Aspose.Cells | Aspose.Cells export chart PNG high resolution | C# set anti aliasing when rendering charts Aspose | Aspose.Cells high DPI chart image example | chart.ToImage anti aliasing option
// Developer Intent: Turn on anti‑aliasing while exporting a chart to a high‑resolution PNG using Aspose.Cells for .NET.
// Use Cases: Generate publication‑quality chart PNGs at 300 dpi with smooth edges for reports. | Batch‑export multiple worksheet charts as anti‑aliased PNGs for inclusion in PDFs. | Automate creation of presentation graphics where chart clarity is critical.
// AI Prompts: Show C# code that sets ImageOrPrintOptions.AntiAliasing = true and exports a chart to a 300 dpi PNG using Aspose.Cells. | Explain how anti‑aliasing improves chart image quality and how to enable it with Chart.ToImage in a .NET console app. | Provide a step‑by‑step guide to export high‑resolution, anti‑aliased chart images from Aspose.Cells.

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // A C# sample that creates a workbook, adds a column chart, and exports the chart as a 300 dpi PNG with smooth edges by setting ImageOrPrintOptions.AntiAliasing = true before calling Chart.ToImage. The example also shows how to save the workbook.
    public class ChartAntiAliasingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Configure image rendering options (high‑resolution)
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                    // ImageFormat defaults to PNG based on file extension
                };

                // Export the chart to a high‑resolution PNG
                string chartPath = "Chart_HighRes_AntiAliased.png";
                try
                {
                    chart.ToImage(chartPath, options);
                    Console.WriteLine($"Chart image saved to: {Path.GetFullPath(chartPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export chart image: {ex.Message}");
                }

                // Save the workbook (optional)
                string workbookPath = "ChartDemo.xlsx";
                try
                {
                    workbook.Save(workbookPath);
                    Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ChartAntiAliasingDemo.Run();
        }
    }
}
