// Title: Export a Waterfall Chart from Excel to PDF with Aspose.Cells for .NET
// Description: This example demonstrates how to load an Excel workbook (WaterfallChart.xlsx), locate the first Waterfall chart, and use Aspose.Cells' Chart.ToPdf method to generate a PDF (WaterfallChart.pdf) where the chart is embedded as an image. The code includes checks for file existence and chart presence, and provides basic error handling.
// Keywords: Aspose.Cells | C# | .NET | Waterfall chart | PDF export | Chart.ToPdf | Excel to PDF conversion | embed chart image | programmatic chart export | example code
// Common Searches: Aspose.Cells export Waterfall chart to PDF C# | Chart.ToPdf example .NET | Convert Excel chart to PDF programmatically | Save specific Excel chart as PDF using Aspose | How to embed Excel chart image in PDF with C#
// Developer Intent: Generate a PDF file that contains only the Waterfall chart extracted from an Excel workbook using Aspose.Cells.
// Use Cases: Produce a PDF report that showcases a single Waterfall chart from a financial model. | Automate batch processing to extract the first chart from multiple workbooks and save each as a separate PDF. | Integrate chart‑to‑PDF conversion into a web API that returns the chart as a downloadable PDF document.
// AI Prompts: Write C# code with Aspose.Cells to open an Excel file, find the first Waterfall chart, and export it to a PDF where the chart is embedded as an image. | Add robust error handling for missing input files and worksheets without charts when converting an Excel chart to PDF. | Explain how the Chart.ToPdf method renders the chart as an image in the PDF and discuss any format or size limitations.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example demonstrates how to load an Excel workbook (WaterfallChart.xlsx), locate the first Waterfall chart, and use Aspose.Cells' Chart.ToPdf method to generate a PDF (WaterfallChart.pdf) where the chart is embedded as an image. The code includes checks for file existence and chart presence, and provides basic error handling.
    public class WaterfallChartToPdf
    {
        public static void Run()
        {
            try
            {
                const string inputFile = "WaterfallChart.xlsx";
                const string outputFile = "WaterfallChart.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the workbook that contains the Waterfall chart
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the worksheet.");
                    return;
                }

                // Assume the Waterfall chart is the first chart in the worksheet
                Chart chart = worksheet.Charts[0];

                // Export the chart directly to a PDF file.
                // The ToPdf method embeds the chart as an image in the resulting PDF.
                chart.ToPdf(outputFile);

                Console.WriteLine($"Waterfall chart has been successfully converted to PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WaterfallChartToPdf.Run();
        }
    }
}
