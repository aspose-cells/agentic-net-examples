// Title: C# – Convert XLSM to PDF with Standard Optimization using Aspose.Cells
// Description: Loads a macro‑enabled XLSM workbook (macros are ignored), configures PdfSaveOptions with OptimizationType = Standard for high‑print‑quality output, and saves the file as a PDF via Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | XLSM to PDF | PdfSaveOptions | Standard optimization | high quality PDF | macro‑enabled workbook conversion | .NET PDF export | US developers | European developers
// Common Searches: Aspose.Cells convert xlsm to pdf c# | PdfSaveOptions Standard optimization example | ignore macros when exporting Excel to PDF with Aspose | C# export macro enabled workbook to high quality PDF | Standard PDF optimization Aspose.Cells .NET
// Developer Intent: Create a PDF from a macro‑enabled XLSM file using Aspose.Cells with Standard (high‑print‑quality) optimization.
// Use Cases: Generate printable PDFs from Excel reports that contain macros, while preserving layout and achieving high print fidelity. | Batch‑process a folder of XLSM templates into PDFs with consistent Standard optimization settings. | Produce PDF invoices or statements from macro‑enabled templates where the macros are not required in the final document.
// AI Prompts: Write C# code that loads an .xlsm file with Aspose.Cells, sets PdfSaveOptions.OptimizationType to Standard, and saves it as a PDF. | Explain the effect of PdfOptimizationType.Standard on PDF quality and file size in Aspose.Cells. | Provide a C# script that iterates through a directory of .xlsm files and converts each to a PDF using Standard optimization.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads a macro‑enabled XLSM workbook (macros are ignored), configures PdfSaveOptions with OptimizationType = Standard for high‑print‑quality output, and saves the file as a PDF via Aspose.Cells for .NET.
class ConvertXlsmToPdfWithStandardOptimization
{
    static void Main()
    {
        // Load the XLSM workbook (macros are ignored during conversion)
        Workbook workbook = new Workbook("input.xlsm");

        // Create PDF save options and set the optimization type to Standard (high print quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
