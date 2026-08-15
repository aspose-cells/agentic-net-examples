// Title: C# – Load Excel with OnlyAuto row auto‑fit and export to PDF using Aspose.Cells
// Description: Demonstrates how to configure LoadOptions.AutoFitterOptions.OnlyAuto to auto‑fit only rows without manual height, then save the workbook as a PDF with formulas calculated, all in a single Aspose.Cells for .NET snippet.
// Keywords: Aspose.Cells | OnlyAuto | AutoFitterOptions | C# | .NET | Excel to PDF | row auto‑fit | LoadOptions | PdfSaveOptions | code example | GitHub demo
// Common Searches: Aspose.Cells OnlyAuto row fitting example | C# load Excel with AutoFitterOptions OnlyAuto | convert Excel to PDF preserving row heights Aspose.Cells | save workbook as PDF after OnlyAuto loading | Aspose.Cells .NET PDF export with auto‑fit rows
// Developer Intent: Load an Excel file with OnlyAuto row auto‑fit enabled and immediately generate a PDF version.
// Use Cases: Create printable PDFs where only automatically sized rows are adjusted, keeping user‑defined heights intact. | Process uploaded spreadsheets in web apps to produce PDFs that respect custom row formatting. | Batch‑convert large sets of Excel reports to PDF while ensuring consistent row‑height handling.
// AI Prompts: Generate C# code that loads an Excel workbook with LoadOptions.AutoFitterOptions.OnlyAuto = true and saves it as a PDF with calculated formulas using Aspose.Cells. | Explain step‑by‑step how OnlyAuto affects row heights during Excel to PDF conversion in Aspose.Cells. | Provide a concise guide for configuring LoadOptions and PdfSaveOptions to export a workbook to PDF while applying OnlyAuto row fitting.

using System;
using Aspose.Cells;

namespace AsposeCellsOnlyAutoFitDemo
{
    // Demonstrates how to configure LoadOptions.AutoFitterOptions.OnlyAuto to auto‑fit only rows without manual height, then save the workbook as a PDF with formulas calculated, all in a single Aspose.Cells for .NET snippet.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Configure load options to enable OnlyAuto auto‑fitting of rows
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AutoFitterOptions = new AutoFitterOptions
            {
                // Fit only rows whose height has not been manually set
                OnlyAuto = true
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Prepare PDF save options (optional settings can be added here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure formulas are calculated before rendering the PDF
                CalculateFormula = true
            };

            // Save the workbook as PDF; rows are already pre‑fitted due to OnlyAuto loading
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook loaded with OnlyAuto row fitting and saved as PDF to '{outputPath}'.");
        }
    }
}
