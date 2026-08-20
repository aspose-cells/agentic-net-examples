// Title: C# – Convert XLSM to PDF with Standard (high‑print) Optimization using Aspose.Cells for .NET
// Description: Loads a macro‑enabled Excel workbook, sets PdfSaveOptions.OptimizationType to Standard for high‑print quality, and saves it as a PDF. Demonstrates Aspose.Cells PDF export with optimal rendering.
// Keywords: Aspose.Cells | C# | XLSM to PDF | PdfSaveOptions | Standard optimization | high print quality | macro-enabled Excel | PDF export .NET | Aspose.Cells PDF optimization | convert Excel to PDF C#
// Common Searches: Aspose.Cells convert XLSM to PDF C# | Standard PDF optimization Aspose.Cells example | How to export macro‑enabled Excel as PDF using .NET | PdfSaveOptions OptimizationType Standard sample code | C# code to save Excel workbook as high‑quality PDF
// Developer Intent: Load a macro‑enabled Excel file, apply the Standard PDF optimization, and generate a PDF document with Aspose.Cells.
// Use Cases: Produce print‑ready PDFs from financial models that contain macros. | Automate batch conversion of XLSM reports to high‑quality PDFs for archival. | Create compliance‑grade PDFs with consistent rendering across platforms. | Integrate PDF export into a .NET web service that receives XLSM uploads.
// AI Prompts: Generate C# code that opens an .xlsm workbook with Aspose.Cells, sets PdfSaveOptions.OptimizationType to Standard, and saves it as a PDF, including error handling. | Show how to combine Standard optimization with other PdfSaveOptions such as ImageQuality, PageCountMode, and embedded fonts. | Write a script that scans a directory, converts each .xlsm file to PDF using Standard optimization, and logs the results.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads a macro‑enabled Excel workbook, sets PdfSaveOptions.OptimizationType to Standard for high‑print quality, and saves it as a PDF. Demonstrates Aspose.Cells PDF export with optimal rendering.
class Program
{
    static void Main()
    {
        // Path to the source XLSM file
        string sourceFile = "input.xlsm";

        // Load the workbook (XLSM format)
        Workbook workbook = new Workbook(sourceFile);

        // Create PDF save options and set the optimization type to Standard (high print quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Path for the resulting PDF file
        string pdfFile = "output.pdf";

        // Save the workbook as PDF using the specified options
        workbook.Save(pdfFile, pdfOptions);
    }
}
