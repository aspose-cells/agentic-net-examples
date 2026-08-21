// Title: C# Load Specific Worksheets (Indexes 0 & 2) and Export to Separate PDFs with Aspose.Cells
// Description: Demonstrates how to create a custom LoadFilter that loads only worksheets at indexes 0 and 2, apply it via LoadOptions, and use PdfSaveOptions with SheetSet to save each loaded sheet as an individual PDF file.
// Keywords: Aspose.Cells | LoadOptions | LoadFilter | specific worksheets | C# PDF export | SheetSet | selective sheet loading | memory optimization | Excel to PDF conversion | index 0 | index 2
// Common Searches: Aspose.Cells load only sheet 0 and 2 | C# export selected Excel sheets to PDF | How to use LoadFilter in Aspose.Cells | Save individual worksheets as PDF with Aspose.Cells | Reduce memory when converting Excel to PDF Aspose
// Developer Intent: Load worksheets 0 and 2 from an Excel workbook and save each as a separate PDF using Aspose.Cells for .NET.
// Use Cases: Generate PDF reports for only the summary and data tabs while skipping other sheets in a large workbook. | Create separate PDF invoices from designated sheets without loading the entire Excel file into memory. | Improve performance in server‑side processing by loading only required worksheets before conversion.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions with a custom LoadFilter to load worksheets at indexes 0 and 2 and saves each to a separate PDF using PdfSaveOptions. | Explain how SheetSet indexes map when only two sheets are loaded and how to reference the second loaded sheet for PDF export. | Show how to modify the example to accept a dynamic list of sheet indexes and generate a PDF for each loaded sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLoadSpecificSheets
{
    // Custom LoadFilter that loads only the sheets with indexes 0 and 2
    // Demonstrates how to create a custom LoadFilter that loads only worksheets at indexes 0 and 2, apply it via LoadOptions, and use PdfSaveOptions with SheetSet to save each loaded sheet as an individual PDF file.
    public class SpecificSheetsLoadFilter : LoadFilter
    {
        // Return the desired sheet indexes; this property is read‑only in the base class
        public override int[] SheetsInLoadingOrder => new int[] { 0, 2 };
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourceFile = "input.xlsx";

            // Configure LoadOptions to use the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new SpecificSheetsLoadFilter();

            // Load the workbook – only sheets 0 and 2 will be present
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Export the first loaded sheet (original index 0) to PDF
            PdfSaveOptions pdfOptions0 = new PdfSaveOptions
            {
                // Restrict saving to the first sheet only
                SheetSet = new SheetSet(new int[] { 0 })
            };
            workbook.Save("Sheet0.pdf", pdfOptions0);

            // Export the second loaded sheet (original index 2) to PDF
            PdfSaveOptions pdfOptions2 = new PdfSaveOptions
            {
                // Restrict saving to the second sheet (which is at index 1 in the loaded workbook)
                // Since only two sheets are loaded, their indexes are 0 and 1 respectively.
                SheetSet = new SheetSet(new int[] { 1 })
            };
            workbook.Save("Sheet2.pdf", pdfOptions2);
        }
    }
}
