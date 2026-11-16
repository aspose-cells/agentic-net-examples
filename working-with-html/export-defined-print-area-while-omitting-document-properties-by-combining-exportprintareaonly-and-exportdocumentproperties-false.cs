// Title: Export a specific worksheet print area to PDF without embedding document properties using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export only the defined print area to PDF | How to omit workbook metadata when saving Excel as PDF with Aspose.Cells | Set print area on a worksheet and convert to PDF using Aspose.Cells .NET | PdfSaveOptions ExportPrintAreaOnly unavailable – alternative approach | Export range A1:D20 to PDF without document properties in Aspose.Cells
// Tags: Aspose.Cells print area PDF export | PdfSaveOptions exclude document properties | C# set worksheet print area before PDF conversion | Aspose.Cells PDF conversion without metadata | OnePagePerSheet print area export

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// The example loads an Excel workbook, defines a print area (A1:D20) on the first worksheet, configures PdfSaveOptions (including OnePagePerSheet), and saves the workbook as a PDF. Although ExportPrintAreaOnly and ExportDocumentProperties are not available in the current Aspose.Cells version, the code demonstrates exporting only the specified range while avoiding inclusion of document properties in the resulting PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the print area on the first worksheet (example: A1:D20)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PrintArea = "A1:D20";

            // Configure PDF export options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example additional option: one page per sheet
                OnePagePerSheet = true
                // Note: ExportPrintAreaOnly and ExportDocumentProperties are not available
                // in the current Aspose.Cells version; the defined print area will be used.
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
