// Title: Export Excel Workbook with Slicers to PDF – Preserve Slicer Visibility & Metadata (C# Aspose.Cells)
// Description: Loads an Excel file, sets every slicer to printable, enables ExportDocumentStructure in PdfSaveOptions, and saves the workbook as a PDF. The resulting PDF retains slicer graphics and workbook metadata.
// Keywords: Aspose.Cells slicer PDF export | C# convert Excel to PDF with slicers | printable slicer Aspose.Cells | ExportDocumentStructure PDF | preserve slicer appearance PDF | Excel workbook metadata PDF
// Common Searches: Aspose.Cells make slicers printable when saving to PDF | C# PDFSaveOptions ExportDocumentStructure example | how to keep slicer graphics in PDF conversion | export Excel workbook with slicers to PDF preserving metadata | Aspose.Cells slicer visibility PDF output
// Developer Intent: Export an Excel workbook that contains slicers to a PDF file while ensuring the slicers appear in the output and the document’s metadata is retained.
// Use Cases: Distribute a financial report with slicer controls as a PDF without losing visual cues. | Create printable dashboards where slicer buttons must be visible in the PDF version. | Archive Excel workbooks with slicers, keeping searchable metadata for compliance.
// AI Prompts: Show C# code to set all slicers printable before saving an Excel file to PDF with Aspose.Cells. | Provide an example that uses PdfSaveOptions.ExportDocumentStructure to retain metadata in the PDF. | Explain how to verify that slicer objects are included in the generated PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an Excel file, sets every slicer to printable, enables ExportDocumentStructure in PdfSaveOptions, and saves the workbook as a PDF. The resulting PDF retains slicer graphics and workbook metadata.
class ConvertWorkbookWithSlicersToPdf
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains slicers
            Workbook workbook = new Workbook(inputPath);

            // Ensure slicers are marked as printable so they appear in the PDF
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    slicer.IsPrintable = true;
                }
            }

            // Configure PDF save options to retain document structure (metadata)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

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
