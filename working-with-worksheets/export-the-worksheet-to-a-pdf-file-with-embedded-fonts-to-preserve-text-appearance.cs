// Title: Export an Excel worksheet to PDF with embedded fonts using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save worksheet as PDF with embedded fonts | How to preserve font appearance when converting Excel to PDF using Aspose.Cells | PdfSaveOptions EmbedStandardWindowsFonts example in .NET | Export Excel file to PDF with all fonts embedded using Aspose.Cells | Aspose.Cells PDF conversion font embedding settings C# code sample
// Tags: Aspose.Cells PDF font embedding | PdfSaveOptions embed fonts | C# Excel to PDF conversion | embed standard windows fonts Aspose | maintain font fidelity PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, configures PdfSaveOptions to embed standard Windows fonts (or all fonts via FontEmbeddingMode when available), and saves the worksheet as a PDF, ensuring the original text appearance is retained while handling missing files and runtime errors.
class ExportWorksheetToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed standard Windows fonts (Arial, Times New Roman, etc.)
                EmbedStandardWindowsFonts = true
                // If the used Aspose.Cells version supports FontEmbeddingMode, you can enable it:
                // FontEmbeddingMode = FontEmbeddingMode.EmbedAll
            };

            // Save the workbook as PDF using the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Export completed successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
