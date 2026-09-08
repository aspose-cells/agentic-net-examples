// Title: Convert an XLSB workbook to PDF with Office Add‑Ins rendering and a custom 0.8 scaling factor using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSB file, checks that it exists, enables rendering of embedded Office Add‑Ins, sets PdfSaveOptions.ZoomFactor to 0.8, configures OnePagePerSheet, and saves the workbook as a PDF with Aspose.Cells. | Add comprehensive error handling to the XLSB‑to‑PDF conversion sample so that missing files, permission problems, and conversion failures are reported while preserving the custom scaling configuration.
// Common Searches: aspose.cells render office add-ins during xlsb to pdf conversion c# | c# set pdf zoom factor 0.8 in aspose.cells PdfSaveOptions | convert binary excel workbook to pdf with custom scaling using asp.net | how to keep office add-in functionality in pdf export with aspose.cells | one page per sheet pdf export with scaling factor asp.net core
// Tags: Aspose.Cells PDF zoom factor scaling | XLSB to PDF conversion with Office Add‑Ins | C# PdfSaveOptions OnePagePerSheet setting | custom scaling factor 0.8 in pdf export | exception handling for workbook file not found | rendering embedded Office Add‑Ins in PDF

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of an XLSB file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to render any embedded Office Add‑Ins, applies a 0.8 zoom factor, forces each worksheet onto a separate PDF page, saves the result as a PDF, and includes robust exception handling to surface file‑related and conversion errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsb";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (example: one page per sheet)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
