// Title: Export an Aspose.Cells workbook to a high‑resolution PDF while preserving crisp slicer text in C#
// AI Prompts: Generate C# code that sets CellsHelper.DPI to 300 and configures PdfSaveOptions with 300 PPI image resampling to create a high‑resolution PDF. | Show how to enable ExportDocumentStructure and use PdfOptimizationType.Standard in PdfSaveOptions to maintain slicer label clarity when saving to PDF. | Provide a complete Aspose.Cells example that creates a worksheet, adds sample data, and saves it as a 300 DPI PDF with optimal print quality.
// Common Searches: Aspose.Cells C# export workbook to PDF with 300 DPI for sharp slicer labels | How to keep slicer text readable when converting Excel to PDF using Aspose.Cells | Set image resample and DPI in PdfSaveOptions for high‑resolution PDF output in C# | Enable document structure in Aspose.Cells PDF export to preserve slicer rendering | Best PDF optimization settings for slicer quality in Aspose.Cells C#
// Tags: Aspose.Cells high‑resolution PDF export | PdfSaveOptions image resample 300 DPI | ExportDocumentStructure Aspose.Cells PDF | slicer text sharpness PDF conversion | standard PDF optimization Aspose.Cells | set CellsHelper DPI for PDF output

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Sets CellsHelper.DPI to 300, configures PdfSaveOptions with 300 PPI image resampling, enables document structure and standard optimization, then saves the workbook as a high‑resolution PDF to keep slicer text sharp.
class ExportWorkbookToPdfHighRes
{
    static void Main()
    {
        // Increase the DPI to obtain higher‑resolution output.
        CellsHelper.DPI = 300;

        // Create a new workbook (or load an existing one).
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data (slicers can be added here if needed).
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Fruits");
        sheet.Cells["A3"].PutValue("Vegetables");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);

        // Configure PDF save options for high‑resolution output.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Resample images to 300 PPI with maximum JPEG quality.
        pdfOptions.SetImageResample(300, 100);

        // Keep document structure (helps with slicer rendering).
        pdfOptions.ExportDocumentStructure = true;

        // Use the standard optimization type for best print quality.
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF with the specified options.
        workbook.Save("output_high_res.pdf", pdfOptions);
    }
}
