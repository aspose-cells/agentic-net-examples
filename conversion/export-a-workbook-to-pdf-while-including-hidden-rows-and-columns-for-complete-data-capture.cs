// Title: Export an Excel workbook to PDF while preserving hidden rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF include hidden rows | How to keep hidden columns when converting XLSX to PDF with Aspose | C# save workbook as PDF with hidden data using PdfSaveOptions | Aspose.Cells PDF conversion retain hidden cells example
// Tags: Aspose.Cells PDFSaveOptions hidden rows | C# workbook save as PDF with hidden cells | Excel to PDF conversion retaining hidden columns | Aspose.Cells export hidden data to PDF | PDF conversion of workbook preserving hidden rows .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program checks for the existence of input.xlsx, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (which includes hidden rows and columns by default), and saves the workbook as output.pdf, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (hidden rows/columns are included by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
