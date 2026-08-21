// Title: C# – Convert XLSX to PDF with Standard‑size optimization using AspNet.Cells
// Description: Loads an XLSX workbook, sets PdfSaveOptions.OptimizationType to Standard (high‑quality size), and saves the workbook as a PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PdfSaveOptions | PdfOptimizationType.Standard | XLSX to PDF C# | .NET Excel to PDF | Standard size PDF | high quality PDF export | Aspose.Cells PDF optimization
// Common Searches: Aspose.Cells set PDF optimization to Standard | C# convert Excel to PDF with high quality | PdfSaveOptions StandardSize example | How to export XLSX as PDF using Aspose.Cells .NET | Standard PDF size Aspose.Cells
// Developer Intent: Export an existing Excel workbook to a PDF with Standard‑size (high‑quality) optimization using Aspose.Cells in C#.
// Use Cases: Generate print‑ready PDF reports from Excel templates while preserving layout fidelity. | Create high‑quality PDF invoices or statements from Excel data for client distribution. | Batch‑process multiple workbooks to PDF with a consistent Standard optimization setting.
// AI Prompts: Write C# code that loads an Excel workbook and saves it as a PDF using PdfOptimizationType.Standard with Aspose.Cells. | Explain the visual differences between Standard, MinimumSize, and MaximumSize PDF optimization types in Aspose.Cells. | Provide best‑practice error handling when converting Excel files to PDF with specific PdfSaveOptions in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an XLSX workbook, sets PdfSaveOptions.OptimizationType to Standard (high‑quality size), and saves the workbook as a PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Create PDF save options and set the optimization type to Standard (high quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF using the specified options
        string outputFile = "output.pdf";
        workbook.Save(outputFile, pdfOptions);
    }
}
