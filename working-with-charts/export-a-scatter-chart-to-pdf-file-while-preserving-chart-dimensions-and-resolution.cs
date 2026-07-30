// Title: Export Scatter Chart to PDF with Exact Size and Resolution using Aspose.Cells for .NET
// Description: Creates a workbook, fills X‑Y data, builds a scatter chart, reads its pixel size, converts to inches via system DPI, and uses Chart.ToPdf to generate a PDF that matches the chart’s original dimensions and centers it on the page.
// Keywords: Aspose.Cells scatter chart PDF export | C# chart to PDF preserving size | convert chart pixels to inches Aspose | Chart.ToPdf exact dimensions | PDF DPI handling Aspose.Cells | .NET export Excel chart as PDF | center chart on PDF page
// Common Searches: How to export a scatter chart to PDF with original size in C# | Aspose.Cells preserve chart resolution when saving to PDF | Convert Excel scatter chart to PDF maintaining dimensions | Chart.ToPdf page size based on pixel dimensions | Center Excel chart on PDF page using Aspose.Cells
// Developer Intent: Generate a PDF file from a scatter chart in an Excel workbook while keeping the chart’s exact pixel dimensions, resolution, and alignment.
// Use Cases: Produce printable reports where the scatter chart must retain its on‑screen size. | Automate batch conversion of multiple Excel scatter charts to PDFs with consistent layout. | Create PDFs for dashboards that require precise chart placement and DPI fidelity.
// AI Prompts: Write C# code with Aspose.Cells to export a scatter chart to PDF, preserving its pixel dimensions and DPI. | Explain how to calculate PDF page width and height from a chart’s pixel size using the system DPI in Aspose.Cells. | Show how to center a chart on a PDF page when exporting with Chart.ToPdf in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsScatterChartPdfExport
{
    // Creates a workbook, fills X‑Y data, builds a scatter chart, reads its pixel size, converts to inches via system DPI, and uses Chart.ToPdf to generate a PDF that matches the chart’s original dimensions and centers it on the page.
    public class ExportScatterChart
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a scatter chart (X values in column A, Y values in column B)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(4);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(6);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(8);
            sheet.Cells["A6"].PutValue(5);
            sheet.Cells["B6"].PutValue(10);

            // Add a scatter chart (X values are categories, Y values are values)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (Y values) and the X values (category data)
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries[0].XValues = "A2:A6";

            // Optional: set chart title
            chart.Title.Text = "Sample Scatter Chart";

            // Ensure the chart uses full page size when exported
            chart.PrintSize = PrintSizeType.Full;

            // Get the actual size of the chart in pixels (returns int[] {width, height})
            int[] actualSize = chart.GetActualSize();
            int widthPixels = actualSize[0];
            int heightPixels = actualSize[1];

            // Get the system DPI (dots per inch)
            double dpi = CellsHelper.DPI; // default is 96

            // Convert pixel dimensions to inches for PDF page size
            float pageWidthInches = (float)(widthPixels / dpi);
            float pageHeightInches = (float)(heightPixels / dpi);

            // Define output file path
            string outputPath = "ScatterChart.pdf";

            // Ensure the directory for the output file exists (handle null when only file name is provided)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // Export the chart to a PDF file, preserving its dimensions and centering it on the page
                chart.ToPdf(outputPath,
                            pageWidthInches,
                            pageHeightInches,
                            PageLayoutAlignmentType.Center,
                            PageLayoutAlignmentType.Center);

                Console.WriteLine("Scatter chart exported to PDF with preserved dimensions.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export chart to PDF: {ex.Message}");
            }
        }
    }
}
