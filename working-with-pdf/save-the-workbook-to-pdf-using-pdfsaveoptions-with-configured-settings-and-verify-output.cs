// Title: Export an Aspose.Cells workbook to PDF with OnePagePerSheet and PDF/A‑1b compliance and validate the generated file in C#
// Common Searches: C# Aspose.Cells export workbook to PDF with PDF/A-1b compliance and one page per sheet | how to verify that a PDF created by Aspose.Cells has content in .NET | using PdfSaveOptions in Aspose.Cells to set PDF/A-1b and single-page-per-sheet | Aspose.Cells save Excel as PDF and check file size programmatically | sample code for saving Aspose.Cells workbook as PDF with compliance settings
// Tags: Aspose.Cells PdfSaveOptions PDF/A-1b compliance | one page per sheet PDF export Aspose.Cells | C# check PDF file existence and size | configure PDF compliance level Aspose.Cells | save Excel workbook as PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with sample data, configures PdfSaveOptions to export each worksheet on a separate page and enforce PDF/A‑1b compliance, saves the workbook as SampleReport.pdf, and then confirms that the PDF file exists and has a non‑zero size.
class PdfExportExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate cells with sample values
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B4"].PutValue(60);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export each worksheet to a separate page
                OnePagePerSheet = true,
                // Set PDF compliance level (e.g., PDF/A-1b)
                Compliance = PdfCompliance.PdfA1b
                // Note: ImageQuality and EmbedStandardFonts are not available in this version of Aspose.Cells
            };

            // Define output PDF path
            string outputPdfPath = Path.Combine(Environment.CurrentDirectory, "SampleReport.pdf");

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPdfPath, pdfOptions);

            // Verify that the PDF file was created and has content
            if (File.Exists(outputPdfPath))
            {
                long fileSize = new FileInfo(outputPdfPath).Length;
                if (fileSize > 0)
                {
                    Console.WriteLine($"PDF successfully created at '{outputPdfPath}'. Size: {fileSize} bytes.");
                }
                else
                {
                    Console.WriteLine("PDF file was created but appears to be empty.");
                }
            }
            else
            {
                Console.WriteLine("Failed to create PDF file.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
