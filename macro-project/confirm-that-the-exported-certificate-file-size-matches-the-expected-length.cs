// Title: C# – Verify the size of a PDF exported from an Excel workbook using Aspose.Cells
// Common Searches: c# how to check the byte size of a PDF created with Aspose.Cells | aspose.cells verify exported PDF file length matches expected value | compare actual PDF size to expected bytes after SaveFormat.Pdf in C# | validate workbook to PDF conversion size using Aspose.Cells API | c# code to throw exception when exported PDF size is incorrect with Aspose.Cells
// Tags: Aspose.Cells PDF export size verification | C# compare exported PDF byte length | Aspose.Cells SaveFormat.Pdf file size check | expected file size validation Aspose.Cells | invalid PDF size exception Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, saves it as a PDF with Aspose.Cells, reads the resulting file's byte length, and optionally compares it to a predefined expected size, throwing an InvalidOperationException when the sizes differ.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Export the workbook to PDF
            string exportedFilePath = "exported.pdf";
            workbook.Save(exportedFilePath, SaveFormat.Pdf);

            // Define the expected file size in bytes (set to 0 to skip verification)
            long expectedFileSize = 0; // TODO: set the expected length if verification is required

            // Get the actual size of the exported file
            long actualFileSize = new FileInfo(exportedFilePath).Length;

            // Verify file size only when an expected size is provided (>0)
            if (expectedFileSize > 0 && actualFileSize != expectedFileSize)
            {
                throw new InvalidOperationException(
                    $"Exported file size mismatch. Expected: {expectedFileSize} bytes, Actual: {actualFileSize} bytes.");
            }

            Console.WriteLine("Export completed successfully.");
            if (expectedFileSize > 0)
            {
                Console.WriteLine("Exported file size matches the expected length.");
            }
            else
            {
                Console.WriteLine($"Exported file size: {actualFileSize} bytes.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
