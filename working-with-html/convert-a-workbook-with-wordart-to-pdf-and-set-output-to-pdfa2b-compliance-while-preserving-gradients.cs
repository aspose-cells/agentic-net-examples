// Title: Convert an Excel workbook containing WordArt to a PDF/A‑2b compliant PDF while preserving gradient fills using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with WordArt, configures PdfSaveOptions for PDF/A‑2b compliance, and saves it as a PDF preserving gradients. | Demonstrate how to verify the input workbook exists before converting it to PDF/A‑2b with Aspose.Cells, ensuring vector graphics are retained. | Show the minimal Aspose.Cells settings required to export a workbook with WordArt to a PDF/A‑2b file without losing gradient colors.
// Common Searches: Aspose.Cells C# convert Excel with WordArt to PDF/A‑2b preserving gradients | How to export workbook containing vector graphics to PDF/A‑2b using .NET | Set PDF/A‑2b compliance in PdfSaveOptions while keeping gradient fills in PDF output
// Tags: Aspose.Cells PDF/A‑2b conversion with WordArt | C# preserve gradient fills in PDF export | PdfSaveOptions compliance setting for PDF/A‑2b | export Excel workbook vector graphics to PDF/A‑2b | WordArt rendering in Aspose.Cells PDF output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook that includes WordArt, sets PdfSaveOptions.Compliance to PdfA2b to retain gradient fills, and saves the file as a PDF/A‑2b document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains WordArt objects
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options for PDF/A‑2b compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the PDF/A compliance level to PDF/A‑2b
                Compliance = PdfCompliance.PdfA2b
                // Aspose.Cells retains gradients and vector graphics by default.
                // No additional settings are required for WordArt rendering.
            };

            // Save the workbook as a PDF/A‑2b compliant PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
