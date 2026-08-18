// Title: C# Aspose.Cells: Convert Workbook to PDF with PDF 1.7 compliance
// Description: This example creates an in‑memory Workbook, adds sample data, configures PdfSaveOptions to use PdfCompliance.Pdf17, and saves the file as a PDF that conforms to the PDF 1.7 standard. The output file (Workbook_V1_7.pdf) demonstrates how to enforce a specific PDF version during export.
// Keywords: Aspose.Cells | C# PDF export | PDF 1.7 compliance | PdfSaveOptions | PdfCompliance.Pdf17 | Excel to PDF conversion | set PDF version | Aspose.Cells PDF save options
// Common Searches: how to set PDF version to 1.7 with Aspose.Cells | Aspose.Cells C# save workbook as PDF 1.7 | PdfCompliance.Pdf17 example code | export Excel workbook to PDF with specific version | C# convert workbook to PDF 1.7 compliance
// Developer Intent: Export a workbook as a PDF that meets PDF 1.7 compliance requirements.
// Use Cases: Generate archival‑grade reports that must follow PDF 1.7 specifications. | Create client‑ready invoices from Excel templates with a guaranteed PDF version. | Automate bulk conversion of Excel files to PDF 1.7 for regulatory submissions.
// AI Prompts: Provide C# code using Aspose.Cells to save a workbook as PDF with PdfCompliance.Pdf17. | Show how to configure PdfSaveOptions for PDF 1.7 compliance in a .NET application. | Explain the steps to change the PDF version when exporting an Excel workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // This example creates an in‑memory Workbook, adds sample data, configures PdfSaveOptions to use PdfCompliance.Pdf17, and saves the file as a PDF that conforms to the PDF 1.7 standard. The output file (Workbook_V1_7.pdf) demonstrates how to enforce a specific PDF version during export.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Version 1.7 Example");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set PDF compliance to version 1.7
            pdfOptions.Compliance = PdfCompliance.Pdf17;

            // Save the workbook as a PDF file with the specified compliance level
            string outputPath = "Workbook_V1_7.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved to PDF with PDF 1.7 compliance at: {outputPath}");
        }
    }
}
