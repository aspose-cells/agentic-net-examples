// Title: Convert an Excel workbook to PDF with Aspose.Cells for .NET and specify PDF version 1.7
// Common Searches: how to specify PDF version 1.7 when saving an Excel file as PDF using Aspose.Cells C# | Aspose.Cells C# example for converting .xlsx to PDF with PDF 1.7 compatibility | set PdfSaveOptions.PdfVersion to 1.7 in Aspose.Cells workbook conversion | C# code to save Excel workbook as PDF 1.7 using Aspose.Cells library | Aspose.Cells PDF conversion with custom PDF version setting in .NET
// Tags: Aspose.Cells PDF version 1.7 | Excel to PDF conversion C# | PdfSaveOptions set PDF compatibility | Workbook.Save PDF output Aspose.Cells | Aspose.Cells .NET PDF conversion example

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an .xlsx file with Aspose.Cells, configuring PdfSaveOptions (including how to set PdfVersion to 1.7), ensuring the output directory exists, and saving the workbook as a PDF. The sample includes basic error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the source workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default PDF version will be used)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Convert and save the workbook as PDF using the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
