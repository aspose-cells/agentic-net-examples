// Title: Export Excel Workbook with Slicers to PDF While Preserving Layout – Aspose.Cells for .NET
// Description: Loads an Excel file that contains slicers, configures PdfSaveOptions (ExportDocumentStructure = true, SheetSet = Visible) and saves the workbook as a PDF so slicer graphics and positions remain unchanged.
// Keywords: Aspose.Cells | C# PDF conversion | .NET Excel to PDF | slicer export | ExportDocumentStructure | SheetSet Visible | preserve slicer layout | dashboard report PDF | Excel slicer rendering
// Common Searches: Aspose.Cells keep slicer formatting when converting to PDF | ExportDocumentStructure option PDFSaveOptions example | Render only visible worksheets with slicers to PDF C# | Preserve slicer appearance in PDF using Aspose.Cells | Convert Excel dashboard with slicers to PDF programmatically
// Developer Intent: Generate a PDF from an Excel workbook that includes slicers, ensuring the slicer visuals and their placement are retained.
// Use Cases: Produce printable PDF versions of interactive dashboards where slicer graphics must match the Excel view. | Automate batch conversion of multiple workbooks that contain slicers, preserving only the visible sheets. | Archive Excel reports with slicer controls as PDFs without losing visual context.
// AI Prompts: Write C# code with Aspose.Cells to save an Excel file containing slicers to PDF while keeping slicer visuals intact. | Explain how ExportDocumentStructure and SheetSet options affect PDF rendering of slicers in Aspose.Cells. | Show error‑handling patterns for converting an Excel workbook with slicers to PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfRender
{
    // Loads an Excel file that contains slicers, configures PdfSaveOptions (ExportDocumentStructure = true, SheetSet = Visible) and saves the workbook as a PDF so slicer graphics and positions remain unchanged.
    class Program
    {
        static void Main()
        {
            // Path to the input Excel file.
            string excelPath = "input_with_slicers.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            try
            {
                // Load the workbook that contains slicers.
                Workbook workbook = new Workbook(excelPath);

                // Configure PDF save options.
                // ExportDocumentStructure = true preserves slicer visual properties and layout.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true,
                    // Render only visible sheets (including slicers).
                    SheetSet = SheetSet.Visible
                };

                // Path for the output PDF file.
                string pdfPath = "output_preserving_slicers.pdf";

                // Save the workbook as a PDF file.
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Workbook successfully rendered to PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
