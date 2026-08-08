// Title: Export Workbook to High‑Resolution PDF with Sharp Slicer Text using Aspose.Cells (C#)
// Description: Sets CellsHelper.DPI to 300, configures PdfSaveOptions (Standard optimization, ExportDocumentStructure, default‑font check) and saves the workbook as a PDF that preserves slicer clarity and print‑ready quality.
// Keywords: Aspose.Cells | C# PDF export | high DPI PDF | CellsHelper.DPI | PdfSaveOptions | OptimizationType.Standard | ExportDocumentStructure | slicer PDF rendering | print quality Excel to PDF | .NET high‑resolution PDF | global | US
// Common Searches: Aspose.Cells increase PDF DPI C# | export slicer to PDF with sharp text | high resolution PDF from Excel Aspose | PdfSaveOptions settings for print quality | how to keep slicer labels clear in PDF
// Developer Intent: Export an Aspose.Cells workbook to a PDF at 300 DPI so that slicer labels and graphics remain crisp and printable.
// Use Cases: Generating printable reports that contain slicers with legible text. | Creating marketing or documentation PDFs where chart and slicer clarity is required. | Batch converting Excel files to high‑resolution PDFs for archival or distribution.
// AI Prompts: Show C# code to export an Aspose.Cells workbook to PDF at 300 DPI while keeping slicer text sharp. | Explain how PdfSaveOptions properties (OptimizationType, ExportDocumentStructure, CheckWorkbookDefaultFont) affect PDF quality for slicers. | Describe the role of CellsHelper.DPI in rendering high‑resolution PDFs with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Sets CellsHelper.DPI to 300, configures PdfSaveOptions (Standard optimization, ExportDocumentStructure, default‑font check) and saves the workbook as a PDF that preserves slicer clarity and print‑ready quality.
    class ExportWorkbookToPdfHighRes
    {
        static void Main()
        {
            // Set the DPI to a high value (e.g., 300) to improve rendering quality of slicer text and other graphics.
            CellsHelper.DPI = 300;

            // Create a new workbook (or load an existing one).
            Workbook workbook = new Workbook();

            // Populate the workbook with some sample data.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruits");
            sheet.Cells["A3"].PutValue("Vegetables");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);

            // (Optional) Add a slicer if needed – omitted here for brevity.

            // Configure PDF save options for high‑quality output.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use standard optimization for best print quality.
                OptimizationType = PdfOptimizationType.Standard,

                // Export document structure (helps retain interactive elements like slicers).
                ExportDocumentStructure = true,

                // Ensure the default font is used if specific fonts are missing.
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as a PDF with the specified options.
            string outputPath = "HighResolutionOutput.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully exported to PDF at high resolution: {outputPath}");
        }
    }
}
