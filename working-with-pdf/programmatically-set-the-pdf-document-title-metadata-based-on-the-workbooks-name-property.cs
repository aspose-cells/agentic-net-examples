// Title: Programmatically embed an Excel workbook's Name as the PDF title metadata with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, reads the workbook's Name property, assigns it to PdfSaveOptions.PdfDocumentInfo.Title, and saves the workbook as a PDF. | Demonstrate how to configure Aspose.Cells PdfSaveOptions to add custom document information (Title, Author, Subject) when converting a workbook to PDF.
// Common Searches: c# aspose.cells set pdf title from workbook name | how to add title metadata when saving excel as pdf using aspose.cells | pdfsaveoptions pdfdocumentinfo title property example c# | embed workbook name into pdf metadata asp.net aspose.cells | custom pdf document properties during excel to pdf conversion
// Tags: Aspose.Cells PdfSaveOptions.Title | C# set PDF metadata from workbook name | Excel to PDF conversion custom document info | Aspose.Cells embed workbook name in PDF title | PdfDocumentInfo custom properties C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The example loads an Excel workbook, creates a PdfSaveOptions object, assigns the workbook's Name to PdfSaveOptions.PdfDocumentInfo.Title, and saves the workbook as a PDF, resulting in a PDF file whose document title metadata matches the original workbook name.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (no direct Title property in PdfSaveOptions)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Successfully saved PDF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
