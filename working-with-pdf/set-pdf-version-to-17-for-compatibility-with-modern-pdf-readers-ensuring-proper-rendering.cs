// Title: Set PDF version to 1.7 when converting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: how to set PDF version 1.7 using Aspose.Cells in C# | Aspose.Cells PdfSaveOptions PdfVersion 1.7 example | C# convert Excel to PDF with specific PDF version Aspose | set PDF compatibility level to 1.7 when saving workbook with Aspose.Cells | Aspose.Cells .NET PDF version 1.7 for modern readers
// Tags: Aspose.Cells set PDF version 1.7 | C# PdfSaveOptions PdfVersion property | Excel workbook to PDF with specific PDF version | Aspose.Cells PDF compatibility level configuration | Save workbook as PDF using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfSaveOptions object, sets its PdfVersion property to PdfVersion.PDF_1_7, and saves the workbook as a PDF, ensuring compatibility with modern PDF readers.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings are used here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
