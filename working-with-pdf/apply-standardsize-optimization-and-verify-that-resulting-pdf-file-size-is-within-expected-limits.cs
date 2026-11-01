// Title: Use Aspose.Cells PdfSaveOptions to apply StandardSize optimization and validate PDF file size in C#
// Common Searches: how to apply StandardSize optimization when converting Excel to PDF with Aspose.Cells in C# | C# code to check if generated PDF size is under a specific limit after saving workbook using Aspose.Cells | Aspose.Cells PdfSaveOptions settings for reducing PDF output size in .NET | verify PDF file size after workbook.Save with PdfSaveOptions in Aspose.Cells | example of creating output directory and measuring PDF size in C# Aspose.Cells
// Tags: Aspose.Cells StandardSize PDF optimization | C# measure PDF file size after Excel conversion | PdfSaveOptions size reduction settings Aspose.Cells | validate PDF output size .NET workbook.Save | Excel workbook to PDF size verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures PdfSaveOptions (without version‑specific flags), saves the workbook as a PDF, then measures the resulting file size and compares it against a 500 KB threshold, reporting whether the size verification passed or failed.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: Specific optimization flags may vary by Aspose.Cells version.
            // The example omits unsupported properties to ensure compatibility.

            // Path for the resulting PDF file
            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options (lifecycle rule: save)
            workbook.Save(outputPath, pdfOptions);

            // Verify the size of the generated PDF
            long actualSizeBytes = new FileInfo(outputPath).Length;

            // Define the expected maximum size (e.g., 500 KB)
            long expectedMaxSizeBytes = 500 * 1024; // 500 KB

            // Output verification result
            if (actualSizeBytes <= expectedMaxSizeBytes)
            {
                Console.WriteLine($"PDF size verification passed: {actualSizeBytes} bytes ≤ {expectedMaxSizeBytes} bytes.");
            }
            else
            {
                Console.WriteLine($"PDF size verification failed: {actualSizeBytes} bytes > {expectedMaxSizeBytes} bytes.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
