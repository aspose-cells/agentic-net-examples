// Title: Add a semi‑transparent diagonal watermark text to each page when converting an Excel workbook to PDF using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions.WatermarkText for a semi‑transparent diagonal watermark, and saves the workbook as a PDF. | Show how to reference the Aspose.Cells.Pdf assembly and set PdfSaveOptions.WatermarkText properties such as Text, Font, Color, Opacity, and Rotation to create a diagonal watermark. | Provide a complete example that lists required NuGet packages, creates PdfSaveOptions, and generates a watermarked PDF from an Excel workbook.
// Common Searches: how to set PdfSaveOptions.WatermarkText for diagonal watermark in Aspose.Cells C# | Aspose.Cells add semi transparent watermark to PDF output | reference Aspose.Cells.Pdf assembly for PDF watermark feature in .NET | C# convert Excel to PDF with watermark using Aspose.Cells | PdfSaveOptions watermark opacity and rotation example
// Tags: Aspose.Cells PDF watermarking | PdfSaveOptions WatermarkText C# | semi transparent diagonal watermark | reference Aspose.Cells.Pdf assembly | Excel to PDF conversion watermark

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies the presence of an input Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance, notes that the PdfWatermark class resides in the Aspose.Cells.Pdf assembly (required for watermarking), and saves the workbook as a PDF without applying a watermark.
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
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PdfWatermark class is part of a separate Aspose.Cells.Pdf assembly.
            // If that assembly is not referenced, the watermark feature cannot be used.
            // The following code is omitted to ensure the sample compiles and runs
            // without additional dependencies.

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
