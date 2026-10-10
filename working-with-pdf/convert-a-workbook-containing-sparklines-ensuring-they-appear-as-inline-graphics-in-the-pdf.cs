// Title: Convert an Excel workbook with sparklines to PDF using Aspose.Cells for .NET, preserving sparklines as inline graphics
// AI Prompts: Write C# code that loads an .xlsx file containing sparklines, configures Aspose.Cells PdfSaveOptions, and saves the workbook as a PDF where the sparklines are rendered as inline images. | Show how to add basic file‑existence checking and exception handling while exporting a workbook with sparklines to PDF with Aspose.Cells.
// Common Searches: asp.net c# export excel file with sparklines to pdf using aspose.cells | how to keep sparklines visible when converting xlsx to pdf with aspose | c# Aspose.Cells PdfSaveOptions default behavior sparklines as images | convert workbook containing sparklines to pdf preserving graphics Aspose.Cells | sample code for exporting sparklines to pdf in .NET
// Tags: aspose.cells pdfsaveoptions sparklines | c# export excel sparklines to pdf | aspose.cells render sparklines as images | convert workbook with sparklines to pdf | asp.net excel to pdf sparklines

using System;
using System.IO;
using Aspose.Cells;

// // Loads an existing .xlsx workbook that includes sparklines, applies default PdfSaveOptions so the sparklines are rendered as inline graphics, and saves the result as a PDF with basic error handling.
class SparklinePdfExport
{
    static void Main()
    {
        const string inputPath = "InputWithSparklines.xlsx";
        const string outputPath = "OutputWithSparklines.pdf";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook that contains sparklines
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default behavior renders sparklines as images)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
