// Title: How to keep hyperlinks active when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# preserve hyperlinks in PDF export | convert Excel workbook to PDF with active links using Aspose.Cells | PdfSaveOptions keep hyperlink functionality when saving as PDF | C# Aspose.Cells PDF conversion retains URL links | enable hyperlink preservation in Aspose.Cells PDF output
// Tags: Aspose.Cells PDFSaveOptions hyperlink retention | Excel to PDF conversion with clickable URLs in .NET | C# preserve active hyperlinks during PDF generation | Aspose.Cells workbook to PDF preserving links | hyperlink support in Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file into an Aspose.Cells Workbook, uses PdfSaveOptions (which retain hyperlinks by default) and saves the workbook as a PDF, ensuring that any hyperlinks in the source spreadsheet remain clickable in the generated PDF.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input Excel file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (hyperlinks are preserved by default)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine("An error occurred while converting the workbook to PDF:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
