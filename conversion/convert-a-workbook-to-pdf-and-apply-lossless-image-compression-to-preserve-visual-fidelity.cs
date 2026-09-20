// Title: Convert an Excel workbook to PDF with lossless image compression using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to preserve image quality, and saves the workbook as a PDF. | Show how to export an Excel workbook to PDF while keeping the original image fidelity by adjusting PdfSaveOptions in a .NET project.
// Common Searches: how to keep image quality when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells PdfSaveOptions lossless image compression example | C# convert .xlsx to PDF without losing visual fidelity using Aspose.Cells
// Tags: Aspose.Cells lossless PDF image compression | C# PdfSaveOptions visual fidelity | Excel to PDF high-quality images Aspose.Cells | Workbook conversion to PDF preserving images .NET | Aspose.Cells PDF export with original image quality

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the existence of input.xlsx, loads it into an Aspose.Cells Workbook, creates PdfSaveOptions (which use lossless image handling by default), and saves the workbook as output.pdf while preserving visual fidelity.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default compression)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
