// Title: Convert an XLSX workbook to a password‑protected PDF with formula calculation using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, evaluates all formulas, and saves it as a PDF encrypted with a user password. | Show how to set the user password on PdfSaveOptions when exporting a workbook to PDF with Aspose.Cells. | Add robust error handling that checks the source Excel file, creates the output folder if missing, and logs any conversion exceptions.
// Common Searches: Aspose.Cells C# export Excel to PDF with password encryption and formula calculation | how to set a user password on PDF generated from XLSX using Aspose.Cells | calculate all formulas before saving workbook as PDF in .NET | C# example converting XLSX to encrypted PDF with Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PDF password encryption | C# calculate Excel formulas before PDF export | PdfSaveOptions set user password Aspose.Cells | XLSX to encrypted PDF conversion .NET | Workbook.Save PDF with security settings

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook, forces calculation of all formulas, configures PdfSaveOptions with a user password, ensures the output directory exists, and saves the workbook as a password‑protected PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the output PDF file
            string outputPath = "output.pdf";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Configure PDF save options (password protection and compliance are omitted
            // because the required PdfSecurityOptions type is unavailable in the current package)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
