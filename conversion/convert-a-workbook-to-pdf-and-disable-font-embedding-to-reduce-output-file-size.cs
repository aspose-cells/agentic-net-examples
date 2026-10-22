// Title: Convert an Excel workbook to PDF without embedding fonts using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert XLSX to PDF without embedding standard fonts | How to reduce PDF size when saving Excel with Aspose.Cells | PdfSaveOptions EmbedStandardWindowsFonts false example in C# | Disable font embedding in Excel to PDF conversion using Aspose.Cells | C# code to save workbook as PDF with no embedded fonts Aspose
// Tags: Aspose.Cells PdfSaveOptions EmbedStandardWindowsFonts false | Excel to PDF conversion without font embedding C# | reduce PDF output size Aspose.Cells | disable standard Windows fonts embedding Aspose | C# workbook to PDF with custom save options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel file (input.xlsx) with Aspose.Cells, configures PdfSaveOptions to set EmbedStandardWindowsFonts to false, and saves the workbook as a PDF (output.pdf). This disables embedding of standard Windows fonts, resulting in a smaller PDF file.
class WorkbookToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to avoid embedding standard Windows fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not embed standard Windows fonts into the PDF to reduce file size
                EmbedStandardWindowsFonts = false
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
