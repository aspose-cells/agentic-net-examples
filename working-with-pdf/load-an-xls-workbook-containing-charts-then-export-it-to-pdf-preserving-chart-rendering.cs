// Title: Export XLS Workbook with Charts to PDF Using Aspose.Cells (.NET)
// Description: Loads an .xls workbook that contains charts, enables RefreshChartCache in PdfSaveOptions to refresh chart data, and saves the workbook as a PDF while preserving the visual appearance of all charts.
// Keywords: Aspose.Cells | .NET | Excel to PDF | XLS charts | RefreshChartCache | PDF conversion | chart rendering | export workbook to PDF | preserve chart images
// Common Searches: Aspose.Cells export XLS with charts to PDF | How to keep Excel chart images when converting to PDF .NET | RefreshChartCache property usage | Convert legacy .xls files to PDF preserving charts | Aspose.Cells PDFSaveOptions chart rendering
// Developer Intent: Convert an Excel .xls workbook that includes charts into a PDF while ensuring the charts render correctly.
// Use Cases: Generate PDF reports from legacy .xls files that contain financial charts. | Batch convert chart‑rich Excel workbooks to PDF for archiving or distribution. | Create PDF invoices or statements where embedded performance charts must appear accurately. | Produce printable PDFs for regulatory submissions that require exact chart fidelity.
// AI Prompts: Provide a C# snippet that loads an .xls file with multiple chart types and saves it as PDF using Aspose.Cells, with RefreshChartCache enabled. | Explain how the RefreshChartCache property influences chart rendering during PDF export and recommend scenarios for its use. | Show how to combine additional PdfSaveOptions such as page orientation, image quality, and chart cache refresh in a single export.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;   // PdfSaveOptions resides in this namespace

// Loads an .xls workbook that contains charts, enables RefreshChartCache in PdfSaveOptions to refresh chart data, and saves the workbook as a PDF while preserving the visual appearance of all charts.
class ExportWorkbookWithChartsToPdf
{
    static void Main()
    {
        // Path to the source Excel file that contains charts
        string sourceFile = "input.xls";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(sourceFile);

        // Configure PDF save options to refresh chart cache so charts render correctly
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            RefreshChartCache = true
        };

        // Save the entire workbook as a PDF file
        string pdfFile = "output.pdf";
        workbook.Save(pdfFile, pdfOptions);

        Console.WriteLine($"Workbook '{sourceFile}' has been exported to PDF as '{pdfFile}'.");
    }
}
