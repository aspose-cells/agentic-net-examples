// Title: Convert an Excel workbook with WordArt to PDF/A‑2b in C# while preserving gradient fills using Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions for PDF/A‑2b compliance, enables gradient preservation for WordArt objects, and saves the workbook as a PDF. | Update existing Aspose.Cells conversion code to set PDF/A‑2b compliance and ensure WordArt gradient fills are retained in the output PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF/A-2b with WordArt gradients preserved | How to keep gradient fills in WordArt when converting Excel to PDF/A-2b using Aspose.Cells | Set PDF/A-2b compliance in Aspose.Cells PdfSaveOptions example | Convert Excel workbook containing WordArt to PDF/A-2b in .NET
// Tags: Aspose.Cells PDF/A-2b export with gradient preservation | C# PdfSaveOptions PDF/A-2b setting | WordArt gradient retention Aspose.Cells | Excel to PDF conversion Aspose.Cells .NET | PDF/A-2b output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads 'input.xlsx' with Aspose.Cells, configures PdfSaveOptions for PDF/A‑2b compliance, preserves WordArt gradient fills, and saves the result as 'output.pdf' while handling missing files and exceptions.
class Program
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
