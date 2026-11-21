// Title: Export an Excel worksheet with embedded OLE objects to PDF while preserving placeholders using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export worksheet containing OLE objects to PDF C# | keep OLE object placeholders when converting Excel to PDF with Aspose.Cells | C# save Excel file as PDF without rendering embedded OLE objects Aspose.Cells | PdfSaveOptions default behavior for OLE placeholders in Aspose.Cells .NET | how to handle embedded OLE objects during Excel to PDF conversion in C#
// Tags: Aspose.Cells PDFSaveOptions OLE placeholders | export Excel worksheet to PDF C# | preserve OLE objects as placeholders Aspose.Cells | convert workbook to PDF with OLE objects Aspose.Cells | C# Aspose.Cells save as PDF OLE handling

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input Excel file, loads it with Aspose.Cells, uses the default PdfSaveOptions (which render OLE objects as placeholders), and saves the workbook or a specific worksheet to a PDF, handling any runtime exceptions.
class ExportOleObjectsToPdf
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains OLE objects.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options. By default, OLE objects are rendered as placeholders.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook (or a specific worksheet) to PDF.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
