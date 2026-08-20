// Title: Hide Gridlines When Converting an XLSM Workbook to PDF with Aspose.Cells for .NET
// Description: Learn how to turn off both visible and printable gridlines in every worksheet, set the PDF GridlineType to Hair, and save an XLSM file as a clean PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells hide gridlines PDF | disable gridlines Aspose.Cells | PdfSaveOptions GridlineType Hair | print gridlines false Aspose.Cells | XLSM to PDF without gridlines | C# Aspose.Cells PDF export | remove solid lines PDF Aspose
// Common Searches: how to hide gridlines when exporting XLSX to PDF Aspose.Cells | Aspose.Cells PDF export without gridlines C# | set GridlineType to Hair Aspose.Cells | disable printable gridlines Aspose.Cells workbook | convert macro enabled workbook to PDF without lines
// Developer Intent: Generate a PDF from an XLSM workbook that contains no solid gridlines.
// Use Cases: Create polished PDF reports from macro‑enabled workbooks for client presentations. | Produce printable invoices or statements where gridlines must be omitted. | Batch‑export multiple worksheets to a single PDF while preserving a clean layout.
// AI Prompts: Show C# code that disables visible and printable gridlines and sets GridlineType to Hair when saving a workbook to PDF with Aspose.Cells. | Explain the effect of IsGridlinesVisible vs. PageSetup.PrintGridlines on PDF output in Aspose.Cells. | Provide a step‑by‑step guide to export an XLSM file to PDF without any gridlines using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

// Learn how to turn off both visible and printable gridlines in every worksheet, set the PDF GridlineType to Hair, and save an XLSM file as a clean PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the XLSM workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Disable visible gridlines for all worksheets
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.IsGridlinesVisible = false;      // hide gridlines in the sheet view
            ws.PageSetup.PrintGridlines = false; // ensure gridlines are not printed to PDF
        }

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set gridline type to a non‑solid style (Hair) – this effectively removes solid lines
            GridlineType = GridlineType.Hair
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
