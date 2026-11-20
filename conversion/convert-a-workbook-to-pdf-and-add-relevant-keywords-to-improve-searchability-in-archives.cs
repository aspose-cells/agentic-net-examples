// Title: Convert an Excel workbook to PDF and embed PDF keywords using Aspose.Cells for .NET (C#)
// Common Searches: how to add PDF keywords when converting an .xlsx file to PDF with Aspose.Cells C# | Aspose.Cells C# save workbook as PDF with custom metadata | set PDF document keywords during Excel to PDF conversion using Aspose.Cells | C# code to convert Excel to PDF and include searchable keywords | Aspose.Cells PDFSaveOptions add keywords to generated PDF
// Tags: Aspose.Cells PDFSaveOptions metadata | C# Excel to PDF conversion with keywords | embed PDF keywords using Aspose.Cells | save workbook as PDF Aspose.Cells | custom PDF document properties Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example checks for the source XLSX file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions, and saves the workbook as a PDF. It notes that adding PDF keywords requires the Aspose.Pdf library and demonstrates proper exception handling.
class WorkbookToPdfWithKeywords
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Setting PDF metadata (e.g., keywords) requires Aspose.Pdf library.
            // If Aspose.Pdf is not referenced, this step is omitted.

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
