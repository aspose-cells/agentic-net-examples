// Title: Convert an Excel workbook with slicers to PDF while preserving slicer visuals and document metadata using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, iterates through every worksheet, sets each slicer's IsPrintable property to true, and saves the workbook as a PDF with ExportDocumentStructure enabled via Aspose.Cells. | Show how to configure PdfSaveOptions in Aspose.Cells to retain slicer appearance and embed the workbook's document structure when converting Excel to PDF.
// Common Searches: C# Aspose.Cells keep slicers visible when exporting Excel to PDF | preserve slicer formatting and metadata during Excel to PDF conversion | how to set slicer IsPrintable property for PDF output with Aspose.Cells | Aspose.Cells PdfSaveOptions ExportDocumentStructure example for slicers
// Tags: Aspose.Cells PDF conversion with slicers | set slicer IsPrintable property C# | export Excel slicers to PDF preserving metadata | PdfSaveOptions ExportDocumentStructure usage | convert workbook with slicers to PDF .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an Excel workbook, marks all slicers as printable, configures PdfSaveOptions to export document structure, ensures the output directory exists, and saves the workbook as a PDF.
class ConvertWorkbookWithSlicersToPdf
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing slicers
            Workbook workbook = new Workbook(inputPath);

            // Ensure all slicers are marked as printable so they appear in the PDF
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                SlicerCollection slicers = sheet.Slicers;
                foreach (Slicer slicer in slicers)
                {
                    slicer.IsPrintable = true;
                }
            }

            // Configure PDF save options to retain document structure (metadata)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
