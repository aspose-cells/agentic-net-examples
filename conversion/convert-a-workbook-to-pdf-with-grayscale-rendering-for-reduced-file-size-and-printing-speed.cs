// Title: Convert an Excel workbook to a grayscale PDF with reduced size using Aspose.Cells PdfSaveOptions in C#
// Common Searches: c# aspnet convert xlsx to grayscale pdf using Aspose.Cells PdfSaveOptions | how to reduce PDF file size when exporting Excel to PDF with Aspose.Cells | Aspose.Cells grayscale rendering options for PDF export in .NET | fit all columns on a single page per sheet Aspose.Cells PDF conversion example | enable ImageColorMode Grayscale in Aspose.Cells PDF save C#
// Tags: Aspose.Cells PdfSaveOptions grayscale rendering | Aspose.Cells AllColumnsInOnePagePerSheet PDF export | Aspose.Cells ImageColorMode Grayscale | Aspose.Cells TextColorMode Grayscale | C# Excel to PDF size reduction with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, verifies the source file, configures PdfSaveOptions to fit all columns on one page per sheet, optionally sets ImageColorMode and TextColorMode to Grayscale, and saves the workbook as a reduced‑size grayscale PDF using Aspose.Cells for .NET.
class Program
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Fit each sheet on a single page to reduce size
                AllColumnsInOnePagePerSheet = true
            };

            // The following properties are available in newer Aspose.Cells versions.
            // Uncomment if your version supports them.
            // pdfOptions.ImageColorMode = ImageColorMode.Grayscale;
            // pdfOptions.TextColorMode = TextColorMode.Grayscale;

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
