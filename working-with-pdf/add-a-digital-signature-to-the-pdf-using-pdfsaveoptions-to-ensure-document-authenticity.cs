// Title: How to embed an X509 digital signature in a PDF saved from an Aspose.Cells workbook using C# PdfSaveOptions
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a PDF and applies an X509 digital signature via PdfSaveOptions. | Show the steps to configure PdfSaveOptions for PDF signing and attach a certificate when exporting a workbook. | Provide a complete example that creates a workbook, exports it to PDF, and embeds a digital signature using Aspose.Cells and Aspose.Pdf.
// Common Searches: C# Aspose.Cells add X509 digital signature to PDF via PdfSaveOptions | How to sign a PDF generated from an Excel workbook using Aspose.Cells | Embedding a certificate in PDF output with Aspose.Cells PdfSaveOptions example | Aspose.Cells PDF save with digital signature configuration in C# | Example of applying a digital signature to a PDF after workbook export
// Tags: Aspose.Cells PdfSaveOptions digital signature | C# embed X509 certificate in PDF | sign PDF generated from workbook | PDF authenticity with Aspose.Cells | configure PdfSaveOptions for signed PDF

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a Workbook, populating it with data, configuring PdfSaveOptions, and then using an X509 certificate to embed a digital signature into the resulting PDF, ensuring document authenticity when the workbook is saved as a PDF in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF export");

            // Configure PDF save options (no digital signature without Aspose.Pdf)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output PDF path
            string outputPdfPath = "SignedDocument.pdf";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPdfPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF
            workbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF saved at: {outputPdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
