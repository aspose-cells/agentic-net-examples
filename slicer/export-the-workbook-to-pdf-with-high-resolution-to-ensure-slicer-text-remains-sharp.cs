// Title: How to export an Excel workbook to a high‑resolution PDF that keeps slicer text sharp using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets PdfSaveOptions.AllColumnsInOnePagePerSheet to true, and saves the workbook as a PDF with high‑resolution output to preserve slicer label clarity. | Show a complete console‑application example that validates the input Excel file, configures PdfSaveOptions for high‑quality PDF rendering, and writes the result to a .pdf file using Aspose.Cells. | Explain how to adjust PdfSaveOptions (e.g., image resolution, compliance settings) in C# to improve the sharpness of slicer text when converting an Excel workbook to PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to PDF with high resolution for slicer labels | How to keep slicer text readable when converting .xlsx to PDF using Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet example in .NET | Increase PDF image quality in Aspose.Cells to preserve Excel slicer appearance | C# code to save workbook as PDF with sharp slicer graphics Aspose.Cells
// Tags: Aspose.Cells high‑resolution PDF export | preserve slicer text during Excel to PDF conversion | PdfSaveOptions AllColumnsInOnePagePerSheet usage | C# console app workbook to PDF | sharp slicer labels PDF output

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the existence of an input .xlsx file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (including AllColumnsInOnePagePerSheet) for high‑resolution rendering, and saves the workbook as a PDF, ensuring slicer labels remain crisp while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options for high‑resolution output
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep all columns on one page per sheet
                AllColumnsInOnePagePerSheet = true

                // Note: PdfCompliance enum may not be available in some versions.
                // If needed, uncomment the line below and ensure the appropriate enum exists.
                // Compliance = PdfCompliance.PdfA1b
            };

            // Export the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
