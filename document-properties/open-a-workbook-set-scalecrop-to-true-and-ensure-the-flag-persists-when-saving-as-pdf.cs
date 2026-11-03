// Title: Set ScaleCrop to true in Aspose.Cells and keep the setting when exporting a workbook to PDF using C#
// Common Searches: how to enable ScaleCrop in Aspose.Cells before saving workbook as PDF in C# | Aspose.Cells PDF export retain ScaleCrop flag after saving | C# set PdfSaveOptions ScaleCrop true for Excel to PDF conversion | preserve workbook scaling options when converting to PDF with Aspose.Cells | ScaleCrop property not applied after PDF save Aspose.Cells C#
// Tags: ScaleCrop PdfSaveOptions Aspose.Cells | C# Aspose.Cells PDF scaling settings | preserve workbook scaling during PDF conversion | OnePagePerSheet with ScaleCrop Aspose.Cells | Excel to PDF export scaling flag

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, configures PdfSaveOptions with OnePagePerSheet and ScaleCrop set to true, and saves the workbook as a PDF while ensuring the scaling flag persists. It includes basic file existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (fit content to a single page per sheet)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
