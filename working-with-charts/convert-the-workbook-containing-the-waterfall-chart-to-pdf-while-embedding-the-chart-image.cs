// Title: Export Waterfall Chart to PDF with Embedded Image using Aspose.Cells for .NET
// Description: Loads an Excel workbook containing a Waterfall chart, exports the first chart to a PNG file, inserts the PNG back into the sheet as a picture, and saves the workbook as a PDF. The resulting PDF shows the chart as an embedded image, preserving its visual layout.
// Keywords: Aspose.Cells | C# | .NET | Waterfall chart | export chart to PDF | chart to image | embed chart in PDF | PdfSaveOptions | Excel to PDF conversion
// Common Searches: Aspose.Cells export waterfall chart to PDF | C# embed Excel chart as image in PDF | convert Excel chart to PNG then PDF | save workbook with chart picture using Aspose.Cells | how to embed chart image in PDF with Aspose.Cells
// Developer Intent: Create a PDF from an Excel file that contains a Waterfall chart, embedding the chart as a raster image.
// Use Cases: Generate PDF reports where chart rendering must match the on‑screen appearance. | Avoid vector‑chart compatibility issues by converting charts to PNG before PDF export. | Validate chart presence before conversion to prevent runtime errors. | Customize PdfSaveOptions (page size, orientation, compression) while embedding chart images.
// AI Prompts: Write C# code that extracts every chart from an Excel workbook, saves each as a PNG, replaces the original charts with the images, and then exports the workbook to a single PDF using Aspose.Cells. | Show how to adjust PdfSaveOptions to set A4 page size, landscape orientation, and image compression when embedding chart images in the PDF. | Explain how to handle workbooks with multiple worksheets and ensure each Waterfall chart is embedded as an image in the final PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace Example
{
    // Loads an Excel workbook containing a Waterfall chart, exports the first chart to a PNG file, inserts the PNG back into the sheet as a picture, and saves the workbook as a PDF. The resulting PDF shows the chart as an embedded image, preserving its visual layout.
    class WaterfallChartToPdf
    {
        static void Main()
        {
            try
            {
                string workbookPath = "WaterfallChart.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input file '{workbookPath}' not found.");
                    return;
                }

                // Load the workbook that already contains a Waterfall chart
                Workbook workbook = new Workbook(workbookPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart in the worksheet
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Assume the Waterfall chart is the first chart in the sheet
                Chart chart = worksheet.Charts[0];

                // Export the chart to an image file (PNG format)
                string chartImagePath = "WaterfallChart.png";
                chart.ToImage(chartImagePath);

                // Insert the exported chart image back into the worksheet as a picture
                // Position it at row 0, column 0 (cell A1)
                worksheet.Pictures.Add(0, 0, chartImagePath);

                // Prepare PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF; the embedded picture will appear in the PDF
                string pdfPath = "WaterfallChart.pdf";
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
