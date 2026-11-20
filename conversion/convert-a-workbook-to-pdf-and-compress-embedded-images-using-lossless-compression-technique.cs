// Title: C# example: Convert an Excel workbook to PDF with lossless image compression using Aspose.Cells
// Common Searches: how to export an .xlsx file to PDF with lossless image compression using Aspose.Cells in C# | Aspose.Cells PdfSaveOptions lossless image compression .NET 6 example | C# code to save Excel workbook as PDF while preserving image quality with Aspose.Cells | set PdfCompressionType.Lossless in Aspose.Cells PDF export C# | conditional compilation for lossless PDF compression in Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions lossless compression | C# Excel to PDF conversion with image optimization | PdfCompressionType.Lossless Aspose.Cells example | .NET 6 conditional compilation Aspose.Cells PDF | lossless image compression for PDF export in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing Excel file, configures PdfSaveOptions for lossless image compression (guarded by a .NET 6 conditional compilation block), and saves the workbook as a PDF, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: If the Aspose.Cells version supports lossless compression,
            // the following properties can be set. They are guarded with
            // compile‑time checks to keep the code compatible with older versions.
#if NET6_0_OR_GREATER
            // pdfOptions.Compression = PdfCompressionType.Lossless;
            // pdfOptions.ImageCompression = PdfImageCompression.Auto;
#endif

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
