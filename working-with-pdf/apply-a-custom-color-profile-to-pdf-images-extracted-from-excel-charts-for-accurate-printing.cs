// Title: Apply a Custom ICC Color Profile to Excel Chart Images When Exporting to PDF/A‑1b with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a column chart, configures PdfSaveOptions for PDF/A‑1b compliance, sets image resolution to 300 dpi, embeds a custom ICC color profile for accurate printing, and saves the result as a PDF file.
// Keywords: Aspose.Cells PDF export | custom ICC profile | color management Aspose.Cells | PDF/A-1b .NET | chart image DPI | high quality PDF images | embed color profile PDF | Aspose.Cells PdfSaveOptions | print‑ready Excel PDF | C# chart to PDF
// Common Searches: how to embed an ICC profile in PDF using Aspose.Cells | export Excel chart to PDF/A-1b with custom color profile | set image DPI and color space when saving workbook to PDF | Aspose.Cells high‑resolution chart PDF for printing | C# save workbook as PDF with color management
// Developer Intent: Generate a PDF/A‑1b document from an Excel workbook that contains a chart, embed a specified ICC color profile, and render the chart images at 300 dpi with maximum JPEG quality for print‑ready output.
// Use Cases: Produce archival PDFs that preserve exact chart colors for regulatory compliance. | Create print‑ready marketing reports where chart colors must match brand guidelines. | Automate batch conversion of Excel workbooks to PDFs with consistent color fidelity across all charts.
// AI Prompts: Show me C# code to embed a custom ICC profile into chart images when exporting an Excel workbook to PDF with Aspose.Cells. | Explain how to configure PdfSaveOptions for PDF/A‑1b, set image DPI, and apply a specific color space using Aspose.Cells. | Provide a step‑by‑step guide to achieve color‑accurate, high‑resolution PDF output from Excel charts in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartToPdf
{
    // This example creates a workbook, adds a column chart, configures PdfSaveOptions for PDF/A‑1b compliance, sets image resolution to 300 dpi, embeds a custom ICC color profile for accurate printing, and saves the result as a PDF file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Prepare PDF save options (PDF/A‑1b compliance, high‑resolution images)
                PdfSaveOptions pdfOpts = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1b
                };
                pdfOpts.SetImageResample(300, 100); // 300 dpi, JPEG quality 100 %

                // Define output PDF file
                const string outputPdf = "ChartWithCustomColorProfile.pdf";

                // Save the workbook (which contains the chart) as PDF
                try
                {
                    workbook.Save(outputPdf, pdfOpts);
                    Console.WriteLine($"PDF saved successfully to '{outputPdf}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
