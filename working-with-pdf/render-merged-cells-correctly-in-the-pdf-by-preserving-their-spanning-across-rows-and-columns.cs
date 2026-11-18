// Title: Preserve merged cells when converting an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# keep merged cells during Excel to PDF conversion | how to export merged ranges to PDF with Aspose.Cells .NET | C# save workbook as PDF preserving merged cell spanning Aspose.Cells | PDF output retains merged cells using Aspose.Cells SaveOptions | Aspose.Cells PdfSaveOptions merged cells not lost in PDF
// Tags: Aspose.Cells PDF export merged cells | PdfSaveOptions preserve merged ranges | C# Excel to PDF conversion merged cells | Aspose.Cells workbook save as PDF | merged cell spanning PDF output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads an Excel file, configures PdfSaveOptions (merged cells are retained by default), ensures the output folder exists, and saves the workbook as a PDF while preserving any merged cell spanning across rows and columns.
class MergedCellsPdfExport
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing merged cells
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (merged cells are preserved by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false
                // Compliance property removed to avoid missing enum reference
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
