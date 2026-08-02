// Title: Export an XLS workbook with charts to PDF using Aspose.Cells for .NET (C#)
// Description: Load an .xls workbook that contains charts, calculate all formulas, enable chart cache refresh, and save the workbook as a PDF while preserving the exact chart rendering.
// Keywords: Aspose.Cells | C# | Export Excel to PDF | Charts to PDF | PdfSaveOptions | RefreshChartCache | CalculateFormula | XLS to PDF conversion | Preserve chart appearance | Aspose.Cells PDF export
// Common Searches: Aspose.Cells export Excel with charts to PDF C# | PdfSaveOptions RefreshChartCache example | Convert .xls file to PDF preserving charts | Calculate formulas before PDF export Aspose.Cells | How to keep chart rendering when saving Excel as PDF
// Developer Intent: Convert an .xls workbook that includes charts into a PDF file, ensuring the charts are rendered accurately.
// Use Cases: Generate PDF versions of financial reports that contain embedded trend charts for stakeholder distribution. | Create PDF invoices that automatically include sales‑performance charts from an existing .xls template. | Automate batch conversion of multiple .xls files with charts to PDFs while maintaining visual fidelity.
// AI Prompts: Write C# code that loads an .xls workbook with charts, runs CalculateFormula, sets PdfSaveOptions.RefreshChartCache to true, and saves the workbook as a PDF using Aspose.Cells. | Explain why the RefreshChartCache option is required for correct chart rendering when converting Excel files to PDF with Aspose.Cells. | Provide a step‑by‑step guide to batch‑process a folder of .xls files containing charts, converting each to PDF and preserving chart quality.

using System;
using Aspose.Cells;

// Load an .xls workbook that contains charts, calculate all formulas, enable chart cache refresh, and save the workbook as a PDF while preserving the exact chart rendering.
class ExportWorkbookWithChartsToPdf
{
    static void Main()
    {
        // Path to the source Excel file that contains charts
        string sourcePath = "input.xls";

        // Load the workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Ensure all formulas are calculated before conversion
        workbook.CalculateFormula();

        // Create PDF save options and enable chart cache refresh
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.RefreshChartCache = true;

        // Export the entire workbook (including charts) to a PDF file
        string pdfPath = "output.pdf";
        workbook.Save(pdfPath, pdfOptions);

        Console.WriteLine("Workbook exported to PDF successfully: " + pdfPath);
    }
}
