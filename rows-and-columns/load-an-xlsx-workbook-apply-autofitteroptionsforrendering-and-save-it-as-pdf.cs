// Title: C# – Load XLSX with AutoFitterOptions.ForRendering and Export to PDF using Aspose.Cells
// Description: Shows how to enable AutoFitterOptions.ForRendering via LoadOptions, load an XLSX workbook, and save it as a PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | AutoFitterOptions | ForRendering | LoadOptions | XLSX to PDF | column auto fit | PDF conversion | Excel rendering | .NET
// Common Searches: Aspose.Cells AutoFitterOptions.ForRendering example | Load XLSX with rendering‑aware auto fit in C# | Convert Excel to PDF preserving column widths Aspose.Cells | How to use LoadOptions with AutoFitterOptions in .NET | C# export XLSX to PDF using Aspose.Cells
// Developer Intent: Load an Excel file with rendering‑aware auto‑fit settings and generate a PDF.
// Use Cases: Create PDF reports that match the on‑screen column widths of the original spreadsheet. | Generate printable invoices from Excel templates while keeping the visual layout intact. | Archive Excel workbooks as PDFs where rendering calculations, not data size, determine column sizing.
// AI Prompts: Provide a C# snippet that sets AutoFitterOptions.ForRendering, loads an XLSX workbook, and saves it as PDF with Aspose.Cells. | Explain the impact of AutoFitterOptions.ForRendering on column widths during Excel‑to‑PDF conversion. | Show how to combine LoadOptions and PdfSaveOptions to export an Excel file to PDF with rendering‑specific auto fitting.

using System;
using Aspose.Cells;

// Shows how to enable AutoFitterOptions.ForRendering via LoadOptions, load an XLSX workbook, and save it as a PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create AutoFitterOptions and enable rendering‑specific fitting
        AutoFitterOptions autoFitOptions = new AutoFitterOptions();
        autoFitOptions.ForRendering = true;

        // Assign the AutoFitterOptions to LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.AutoFitterOptions = autoFitOptions;

        // Load the XLSX workbook using the configured LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Create default PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
