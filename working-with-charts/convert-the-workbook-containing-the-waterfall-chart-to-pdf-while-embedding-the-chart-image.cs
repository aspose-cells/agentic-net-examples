// Title: Convert an Excel workbook with a Waterfall chart to PDF while embedding the chart using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert Excel file containing a waterfall chart to PDF | How to embed a Waterfall chart in a PDF generated from an Excel workbook using Aspose.Cells | Aspose.Cells export chart to PNG then save workbook as PDF with OnePagePerSheet option | Save Excel worksheet as PDF with chart image retained using PdfSaveOptions in C# | Render Waterfall chart to an image stream with Aspose.Cells before PDF conversion
// Tags: waterfall chart PDF conversion Aspose.Cells | chart image export PNG Aspose.Cells | PdfSaveOptions OnePagePerSheet usage | chart.ToImage rendering to memory stream | Excel to PDF with embedded charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace WaterfallChartToPdf
{
    // The example loads a workbook that contains a Waterfall chart, optionally renders the first chart to a PNG image using ImageOrPrintOptions and Chart.ToImage, configures PdfSaveOptions (OnePagePerSheet) to place each worksheet on its own PDF page, and saves the workbook as a PDF where the chart is automatically embedded.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "WaterfallChart.xlsx";
            string outputPdfPath = "WaterfallChart.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that contains the Waterfall chart
                Workbook workbook = new Workbook(inputPath);

                // OPTIONAL: Export the first chart to an image file (PNG) for verification
                ExportFirstChartAsImage(workbook, Path.ChangeExtension(outputPdfPath, ".png"));

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Render each worksheet on a separate PDF page.
                    OnePagePerSheet = true
                };

                // Save the workbook as PDF. The Waterfall chart will be embedded automatically.
                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPdfPath}\"");
                Console.WriteLine($"Chart image exported for reference: \"{Path.ChangeExtension(outputPdfPath, ".png")}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }

        /// <param name="workbook">The workbook containing the chart.</param>
        /// <param name="imagePath">The file path where the PNG image will be saved.</param>
        private static void ExportFirstChartAsImage(Workbook workbook, string imagePath)
        {
            try
            {
                // Iterate through worksheets to locate the first chart.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Charts.Count > 0)
                    {
                        // Get the first chart (assumed to be the Waterfall chart).
                        Chart chart = sheet.Charts[0];

                        // Set image rendering options (default format is PNG).
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                        {
                            HorizontalResolution = 150,
                            VerticalResolution = 150
                        };

                        // Render the chart to an image stream.
                        using (MemoryStream imgStream = new MemoryStream())
                        {
                            chart.ToImage(imgStream, imgOptions);
                            File.WriteAllBytes(imagePath, imgStream.ToArray());
                        }

                        Console.WriteLine($"Chart exported as image: \"{imagePath}\"");
                        return;
                    }
                }

                Console.WriteLine("No chart found in the workbook to export as an image.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to export chart as image:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
