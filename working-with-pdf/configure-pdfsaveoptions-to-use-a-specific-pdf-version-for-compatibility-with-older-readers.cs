// Title: Configure Aspose.Cells PdfSaveOptions for PDF 1.4 compatibility in C#
// Description: This C# example creates a workbook, adds sample data, sets PdfSaveOptions.Compliance to PdfCompliance.Pdf14, and saves the file as Output_Pdf14.pdf, ensuring the generated PDF conforms to version 1.4 for legacy reader support.
// Keywords: Aspose.Cells | PdfSaveOptions | PDF 1.4 | PdfCompliance.Pdf14 | C# PDF export | legacy PDF readers | Excel to PDF conversion | PDF version control
// Common Searches: Aspose.Cells set PDF version C# | PdfSaveOptions compliance PDF 1.4 example | Export Excel as PDF 1.4 using Aspose.Cells | How to force PDF 1.4 output in .NET
// Developer Intent: Generate a PDF from an Excel workbook that complies with PDF 1.4 to guarantee compatibility with older PDF viewers.
// Use Cases: Delivering reports that must open in legacy PDF readers. | Archiving Excel data in a PDF format required by older standards. | Meeting regulatory or contractual mandates that specify PDF 1.4.
// AI Prompts: Show how to set PdfSaveOptions.Compliance to Pdf13 in Aspose.Cells (C#). | Provide a C# snippet that saves a workbook as PDF/A‑1b using Aspose.Cells. | Explain the differences between PdfCompliance enum values in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfVersionDemo
{
    // This C# example creates a workbook, adds sample data, sets PdfSaveOptions.Compliance to PdfCompliance.Pdf14, and saves the file as Output_Pdf14.pdf, ensuring the generated PDF conforms to version 1.4 for legacy reader support.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF version compatibility demo");

            // Create PDF save options (lifecycle: create)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the PDF compliance level to PDF 1.4 for older readers
            pdfOptions.Compliance = PdfCompliance.Pdf14;

            // Save the workbook as PDF with the specified compliance (lifecycle: save)
            workbook.Save("Output_Pdf14.pdf", pdfOptions);

            Console.WriteLine("PDF saved with PDF 1.4 compliance.");
        }
    }
}
