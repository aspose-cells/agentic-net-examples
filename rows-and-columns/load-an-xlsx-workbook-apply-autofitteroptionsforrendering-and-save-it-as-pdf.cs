// Title: Load an XLSX workbook with AutoFitterOptions.ForRendering and save it as PDF using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to configure AutoFitterOptions.ForRendering in LoadOptions, load an XLSX file with those settings, and export the workbook to a PDF using default PdfSaveOptions. This ensures column widths are rendered exactly as they appear on screen.
// Keywords: Aspose.Cells | AutoFitterOptions | ForRendering | LoadOptions | C# | .NET | XLSX to PDF conversion | column auto fit rendering | Excel PDF export | preserve column width
// Common Searches: Aspose.Cells load XLSX with AutoFitterOptions.ForRendering | C# convert Excel to PDF keeping column widths | How to enable rendering auto‑fit in Aspose.Cells | Export XLSX to PDF with exact layout Aspose.Cells | AutoFitterOptions ForRendering example
// Developer Intent: Apply rendering‑specific auto‑fit when loading an Excel workbook and then generate a PDF that matches the on‑screen layout.
// Use Cases: Create PDF reports from Excel templates where precise column widths are required. | Batch‑process spreadsheets for archival PDFs without losing visual fidelity. | Generate printable documents from user‑edited Excel files while preserving the view seen in the application.
// AI Prompts: Provide C# code that loads an XLSX file with AutoFitterOptions.ForRendering enabled and saves it as a PDF using Aspose.Cells. | Explain the impact of AutoFitterOptions.ForRendering on PDF output when converting Excel files with Aspose.Cells. | Show how to configure LoadOptions with rendering‑specific auto‑fit for accurate column widths in PDF conversion.

using System;
using Aspose.Cells;

// Demonstrates how to configure AutoFitterOptions.ForRendering in LoadOptions, load an XLSX file with those settings, and export the workbook to a PDF using default PdfSaveOptions. This ensures column widths are rendered exactly as they appear on screen.
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
