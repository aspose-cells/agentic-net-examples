// Title: Set PDF version to 1.6 (PdfCompliance.Pdf16) using Aspose.Cells for .NET
// Description: Shows how to create or load a workbook, set PdfSaveOptions.Compliance to PdfCompliance.Pdf16, and save the workbook as a PDF that conforms to PDF 1.6, unlocking features such as transparency and layered graphics.
// Keywords: Aspose.Cells PDF 1.6 | PdfSaveOptions Compliance | PdfCompliance.Pdf16 | PDF transparency Aspose.Cells | export workbook to PDF 1.6 | C# Aspose.Cells PDF version | advanced PDF features .NET
// Common Searches: Aspose.Cells set PDF version 1.6 | PdfSaveOptions Compliance property example | enable transparency in PDF with Aspose.Cells | supported PDF compliance levels Aspose.Cells | C# save workbook as PDF 1.6
// Developer Intent: Configure the PDF compliance level to 1.6 so the generated file supports advanced PDF capabilities.
// Use Cases: Export charts with transparent backgrounds to a PDF that retains the transparency effect. | Create invoices or reports that require PDF 1.6 features such as layered graphics or opacity control. | Generate printable documents with complex visual effects by saving the workbook with PdfCompliance.Pdf16.
// AI Prompts: How can I change the code to use PDF 1.7 compliance instead of 1.6? | What additional PDF features become available when using PdfCompliance.Pdf16 in Aspose.Cells? | Provide a C# example that adds a semi‑transparent shape to a worksheet before saving it as a PDF 1.6 document.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfVersionDemo
{
    // Shows how to create or load a workbook, set PdfSaveOptions.Compliance to PdfCompliance.Pdf16, and save the workbook as a PDF that conforms to PDF 1.6, unlocking features such as transparency and layered graphics.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF version 1.6 demo with Aspose.Cells");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the PDF compliance level to PDF 1.6 to enable advanced features (e.g., transparency)
            pdfOptions.Compliance = PdfCompliance.Pdf16;

            // Save the workbook as a PDF file with the specified compliance level
            workbook.Save("PdfVersion16Demo.pdf", pdfOptions);

            Console.WriteLine("PDF generated with version 1.6 compliance.");
        }
    }
}
