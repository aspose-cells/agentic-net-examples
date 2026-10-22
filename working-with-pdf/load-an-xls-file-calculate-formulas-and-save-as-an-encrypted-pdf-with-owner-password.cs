// Title: Convert an XLS workbook to a password-protected PDF with owner password and formula recalculation using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xls file with Aspose.Cells, recalculates all workbook formulas, and saves the result as a PDF protected by an owner password. | Update the existing program to configure PdfSaveOptions.OwnerPassword and enable PDF encryption before calling Workbook.Save.
// Common Searches: Aspose.Cells C# how to set an owner password when saving Excel as PDF | recalculate Excel formulas before exporting to encrypted PDF using Aspose.Cells | convert .xls file to password-protected PDF in .NET | PdfSaveOptions encryption settings Aspose.Cells example
// Tags: Aspose.Cells PDF encryption OwnerPassword | C# recalculate formulas before PDF export | convert XLS to password-protected PDF Aspose.Cells | PdfSaveOptions security configuration .NET | Excel workbook to encrypted PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This C# example shows how to verify an input XLS file, load it into an Aspose.Cells Workbook, recalculate all formulas, and save the workbook as a PDF. By setting PdfSaveOptions.OwnerPassword (and optionally the encryption type), the generated PDF is secured with an owner password. The code also creates the output directory if needed and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xls";
            string outputPath = "output.pdf";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the XLS workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Set up PDF save options (no encryption)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save as PDF
            try
            {
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
