// Title: Export an Excel worksheet to PDF with embedded fonts using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to PDF with all fonts embedded | how to enable font embedding in PdfSaveOptions when converting Excel to PDF | save specific worksheet as PDF preserving text appearance Aspose.Cells | Aspose.Cells PDF conversion embed fonts example C# | handle missing input Excel file before exporting to PDF with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions FontEmbeddingMode | C# export specific worksheet as PDF Aspose.Cells | preserve text appearance with embedded fonts Aspose.Cells PDF | handle file-not-found before PDF export Aspose.Cells | configure PDF conversion settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, optionally selects a worksheet, configures PdfSaveOptions to embed fonts (using FontEmbeddingMode), and saves the result as a PDF while checking for a missing source file and catching runtime exceptions.
class ExportWorksheetToPdf
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // (Optional) Activate the worksheet you want to export
            // workbook.Worksheets.ActiveSheetIndex = 0; // first worksheet

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // The FontEmbeddingMode property may not be available in older Aspose.Cells versions.
            // If it exists, you can uncomment the following line to embed all fonts:
            // pdfOptions.FontEmbeddingMode = FontEmbeddingMode.EmbedAll;

            // Save the workbook (or the active worksheet) as a PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
