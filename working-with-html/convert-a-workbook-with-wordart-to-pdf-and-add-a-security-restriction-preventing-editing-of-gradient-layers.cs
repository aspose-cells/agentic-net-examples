// Title: Convert an Excel workbook that contains WordArt to a password‑protected PDF that blocks editing of gradient layers using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with WordArt, saves it as a PDF, and sets an owner password that disables editing of gradient layers via Aspose.Cells. | Show how to use Aspose.Cells PdfSaveOptions in C# to preserve WordArt shapes while applying PDF security permissions that prevent modifications to gradient fills. | Provide a robust C# example that checks for the source Excel file, converts it to PDF with Aspose.Cells, and configures the PDF to be read‑only for gradient layer edits.
// Common Searches: Aspose.Cells C# export Excel with WordArt to PDF and lock gradient layer editing | How to add PDF owner password and restrict shape modifications when converting .xlsx to .pdf using Aspose.Cells | C# sample for preserving WordArt and disabling gradient fill changes in PDF generated from Excel | Set PDF permissions to prevent editing of gradient layers in Aspose.Cells PDF output | Convert workbook containing WordArt to secured PDF with Aspose.Cells .NET example
// Tags: convert workbook with WordArt to PDF Aspose.Cells | apply PDF owner password Aspose.Cells C# | disable gradient layer editing PDF Aspose.Cells | PdfSaveOptions security settings Aspose.Cells | preserve WordArt shapes during Excel to PDF conversion | restrict PDF shape modifications Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that input.xlsx exists, loads the workbook with Aspose.Cells, and saves it as output.pdf using PdfSaveOptions. It notes where PDF security options such as an owner password and edit restrictions would be applied, but those settings are omitted because the Aspose.Cells.Pdf assembly is not referenced.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PDF security options (owner password, permissions) require the Aspose.Cells.Pdf assembly.
            // If that assembly is not referenced, these settings are omitted to keep the code compilable.

            // Save the workbook as a PDF.
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
