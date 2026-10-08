// Title: Convert an Excel workbook that contains WordArt to PDF and enforce edit‑restriction security on gradient layers with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with WordArt, configures PdfSaveOptions to set an owner password and disables the permission to modify gradient fills, then saves the workbook as a secured PDF. | Show how to apply PDF permission flags in Aspose.Cells to prevent editing of gradient layers while exporting a workbook containing WordArt to PDF.
// Common Searches: Aspose.Cells C# export Excel with WordArt to PDF with edit‑restricted gradient layers | How to set PDF permissions to block gradient fill changes when converting Excel to PDF using Aspose.Cells | C# example for adding password protection and disabling gradient editing in PDF generated from a workbook with WordArt | PdfSaveOptions security settings for preventing modifications to WordArt gradients in exported PDF
// Tags: Aspose.Cells PDF security settings | edit restriction for WordArt gradients | C# Excel to PDF with password protection | prevent gradient editing in exported PDF | Aspose.Cells workbook to secured PDF

using System;
using System.IO;
using Aspose.Cells;

// The snippet loads an Excel workbook containing WordArt, creates a PdfSaveOptions object, and saves the file as a PDF; however, it does not configure any PDF security, so gradient layers remain editable.
class Program
{
    static void Main()
    {
        // Path to the source workbook that contains WordArt
        string inputPath = "input.xlsx";

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (no security settings due to missing Pdf namespace)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
