// Title: C# – Convert XLSX to PDF with Standard‑size Optimization using Aspose.Cells
// Description: Load an XLSX workbook, configure PdfSaveOptions.OptimizationType to Standard for high‑print‑quality output, and save the workbook as a PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF conversion | PdfSaveOptions Standard optimization | C# convert XLSX to PDF | high quality PDF Aspose.Cells | Excel to PDF Aspose .NET
// Common Searches: Aspose.Cells set PdfOptimizationType Standard | C# convert Excel to PDF with high quality | Save workbook as PDF using PdfSaveOptions | Standard size PDF output Aspose.Cells | How to export XLSX to PDF in .NET
// Developer Intent: Generate a PDF from an Excel workbook using Aspose.Cells with Standard‑size (high‑quality) optimization.
// Use Cases: Produce print‑ready PDFs of financial statements while preserving layout and graphics. | Batch‑convert a directory of XLSX files to high‑quality PDFs for archival storage. | Export Excel dashboards to PDF for client distribution, ensuring charts retain original fidelity.
// AI Prompts: Write C# code that iterates over all .xlsx files in a folder and saves each as a PDF using PdfOptimizationType.Standard with Aspose.Cells. | Compare PdfOptimizationType.Standard and MinimumSize in terms of file size, rendering speed, and visual quality. | Show how to combine PdfSaveOptions.Standard with page orientation and margin settings in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an XLSX workbook, configure PdfSaveOptions.OptimizationType to Standard for high‑print‑quality output, and save the workbook as a PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options and set the optimization type to Standard (high print quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
