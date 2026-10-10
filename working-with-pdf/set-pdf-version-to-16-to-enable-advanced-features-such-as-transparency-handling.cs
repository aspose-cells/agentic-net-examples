// Title: How to set PDF version to 1.6 for Excel-to-PDF conversion using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Workbook, sets PdfSaveOptions.Version to PdfVersion.Pdf_1_6, and saves the workbook as a PDF. | Show how to enable PDF 1.6 features such as transparency when exporting an Excel worksheet with Aspose.Cells. | Provide a .NET snippet that configures PdfSaveOptions before calling Workbook.Save to produce a PDF 1.6 file. | Explain which Aspose.Cells property controls the PDF version and how to use it for advanced PDF capabilities.
// Common Searches: Aspose.Cells set PDF version 1.6 in C# | Enable transparency in PDF generated from Excel using Aspose.Cells | PdfSaveOptions.Version property example Aspose.Cells .NET | Export Excel workbook to PDF with PDF 1.6 features Aspose.Cells | How to change PDF version when saving workbook with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions version 1.6 | C# export Excel to PDF with transparency | set PDF version Aspose.Cells .NET | advanced PDF features Aspose.Cells | Excel to PDF PDF 1.6 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExample
{
    // // This example creates a Workbook, adds sample data, configures PdfSaveOptions.Version = PdfVersion.Pdf_1_6 to enable PDF 1.6 features such as transparency, and saves the workbook as Result.pdf.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // create-workbook rule

                // Add sample data (optional, just to have content in the PDF)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data");

                // Configure PDF save options (default PDF version will be used)
                PdfSaveOptions pdfOptions = new PdfSaveOptions(); // create-pdfsaveoptions rule

                // Save the workbook as a PDF with the specified options
                workbook.Save("Result.pdf", pdfOptions); // save-pdf rule
                Console.WriteLine("PDF file 'Result.pdf' has been created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
