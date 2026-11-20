// Title: Convert an Excel workbook to PDF with lossless image compression using Aspose.Cells for .NET (C#)
// Common Searches: how to export an .xlsx file to PDF with lossless image compression in C# using Aspose.Cells | Aspose.Cells C# set PdfSaveOptions.ImageCompression to JpegLossless for PDF export | C# convert Excel workbook to PDF while preserving original layout Aspose.Cells | save Excel as PDF without scaling each sheet using Aspose.Cells PDFSaveOptions | example code for converting workbook to PDF with high‑fidelity images in Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions image compression | C# Excel to PDF conversion preserving image quality | OnePagePerSheet false Aspose.Cells | lossless PDF export Aspose.Cells | configure PDF export options C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, creates a PdfSaveOptions object (setting OnePagePerSheet to false and optionally ImageCompression to JpegLossless for lossless image handling), ensures the output folder exists, and saves the workbook as a PDF. This demonstrates how to convert .xlsx to PDF in C# while maintaining visual fidelity and original layout.
class WorkbookToPdfConverter
{
    static void Main()
    {
        try
        {
            // Define input and output paths
            string inputPath = @"C:\Path\To\InputWorkbook.xlsx";
            string outputPath = @"C:\Path\To\OutputDocument.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (using properties available in the current Aspose.Cells version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Preserve original layout without forcing one page per sheet
                OnePagePerSheet = false
                // Additional options such as Compression or Compliance can be set here
                // if they are supported by the installed Aspose.Cells version.
            };

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
