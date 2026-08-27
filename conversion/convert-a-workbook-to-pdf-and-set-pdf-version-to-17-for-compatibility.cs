// Title: Convert an Aspose.Cells workbook to PDF with PDF 1.7 compliance using C#
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells and saves it as a PDF, configuring PdfSaveOptions.Compliance to Pdf17. | Demonstrate how to set the PDF compliance level to 1.7 before calling Workbook.Save in an Aspose.Cells application.
// Common Searches: asp.net aspose.cells export workbook to pdf with pdf 1.7 compliance | c# set pdf version 1.7 when converting excel to pdf using aspose.cells | how to use PdfSaveOptions to specify PDF 1.7 compliance in Aspose.Cells | pdfsaveoptions compliance property example for Aspose.Cells C# | save excel as pdf version 1.7 with aspose.cells library
// Tags: Aspose.Cells PDF conversion compliance setting | PdfSaveOptions compliance property C# | Workbook.Save PDF version control Aspose.Cells | Export Excel workbook to PDF 1.7 Aspose | Set PDF 1.7 compliance in Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // The program creates a workbook, adds sample data, configures PdfSaveOptions.Compliance to Pdf17, and saves the workbook as a PDF file named Workbook_V1_7.pdf.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (empty or you can load an existing file)
            Workbook workbook = new Workbook();

            // Add some sample data to demonstrate the conversion
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF version 1.7 example");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set PDF compliance to PDF 1.7 (Pdf17)
            pdfOptions.Compliance = PdfCompliance.Pdf17;

            // Save the workbook as PDF with the specified compliance level
            string outputPath = "Workbook_V1_7.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF with version 1.7 to: {outputPath}");
        }
    }
}
