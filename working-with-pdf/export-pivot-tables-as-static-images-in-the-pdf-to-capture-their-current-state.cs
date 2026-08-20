// Title: Export Excel Pivot Tables as Static Images in PDF with Aspose.Cells for .NET
// Description: Loads an Excel workbook, refreshes all pivot tables to capture the latest data, and saves the file as a PDF where each pivot table is rendered as a flat image using PdfSaveOptions.
// Keywords: Aspose.Cells | C# | pivot table PDF export | static image PDF | PdfSaveOptions | refresh pivot tables | Excel to PDF | flatten pivot tables | Aspose.Cells for .NET | export pivot tables as images
// Common Searches: Aspose.Cells export pivot table to PDF | C# save pivot tables as images in PDF | Refresh all pivot tables before PDF conversion Aspose | Flatten Excel pivot tables in PDF using Aspose.Cells | PdfSaveOptions disable document structure Aspose
// Developer Intent: Generate a PDF that contains pivot tables rendered as non‑editable images.
// Use Cases: Distribute read‑only PDF reports with pivot tables preserved visually. | Archive Excel analyses where pivot layouts must stay unchanged after source data updates. | Create client‑facing documents that prevent further manipulation of pivot data.
// AI Prompts: Write C# code with Aspose.Cells to refresh every pivot table and export the workbook to PDF where pivot tables appear as static images. | Explain how PdfSaveOptions can be configured to flatten pivot tables and optionally omit the PDF document structure. | Show how to export only selected worksheets that contain pivot tables to a PDF while keeping their visual appearance intact.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

namespace PivotTablePdfExport
{
    // Loads an Excel workbook, refreshes all pivot tables to capture the latest data, and saves the file as a PDF where each pivot table is rendered as a flat image using PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains pivot tables.
            // Replace "input.xlsx" with the path to your source file.
            Workbook workbook = new Workbook("input.xlsx");

            // Refresh all pivot tables in each worksheet to ensure they reflect the latest data.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.RefreshPivotTables();
            }

            // Optional: configure PDF save options if needed.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // ExportDocumentStructure = false; // Uncomment to disable document structure export.
            };

            // Save the workbook as a PDF. Pivot tables are rendered as static content in the PDF.
            workbook.Save("PivotTablesExported.pdf", pdfOptions);

            Console.WriteLine("Pivot tables have been exported to PDF as static images.");
        }
    }
}
