// Title: Batch export worksheets to individual PDFs with auto‑fitted rows using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, iterates through all worksheets, auto‑fits each row to its content, configures PdfSaveOptions (OnePagePerSheet) and SheetSet to target a single sheet, then saves each worksheet as a separate PDF (e.g., Sheet1.pdf). The process logs each conversion to the console.
// Keywords: Aspose.Cells | C# | .NET | auto fit rows | PDF export | batch worksheet conversion | OnePagePerSheet | SheetSet | save each sheet as PDF | Excel to PDF
// Common Searches: Aspose.Cells auto fit rows before PDF export | C# batch convert Excel worksheets to separate PDFs | How to use SheetSet with PdfSaveOptions in Aspose.Cells | Export each Excel sheet as one‑page PDF using Aspose.Cells .NET | Auto‑fit rows for PDF rendering Aspose.Cells
// Developer Intent: Automatically adjust row heights in every worksheet and generate a separate PDF file for each sheet.
// Use Cases: Create individual PDF reports for each tab of a financial workbook after row height optimization. | Build a command‑line tool that converts large multi‑sheet Excel files into one‑page‑per‑sheet PDFs. | Automate document generation where each worksheet represents a distinct form that must be printed as a PDF with correct layout. | Integrate into a server‑side service that receives an Excel file and returns a zip of per‑sheet PDFs.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, auto‑fits rows, and saves each worksheet as a separate PDF using OnePagePerSheet. | Show how to modify the example to also auto‑fit columns and name PDF files after the worksheet names. | Add robust error handling and logging to the batch PDF export routine. | Explain the purpose of SheetSet in PdfSaveOptions and how to export a range of sheets. | Provide a PowerShell script that calls the compiled .NET executable to process files in a folder.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchWorksheetPdfExport
{
    // Loads an Excel workbook, iterates through all worksheets, auto‑fits each row to its content, configures PdfSaveOptions (OnePagePerSheet) and SheetSet to target a single sheet, then saves each worksheet as a separate PDF (e.g., Sheet1.pdf). The process logs each conversion to the console.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Auto‑fit all rows in the current worksheet to ensure proper rendering
                sheet.AutoFitRows();

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Render each sheet on a single page (optional, adjust as needed)
                    OnePagePerSheet = true
                };

                // Restrict the save operation to the current sheet only
                pdfOptions.SheetSet = new SheetSet(new int[] { i });

                // Define the output PDF file name for the current sheet
                string outputFile = $"Sheet{i + 1}.pdf";

                // Save the current worksheet as a PDF using the configured options
                workbook.Save(outputFile, pdfOptions);

                Console.WriteLine($"Saved worksheet '{sheet.Name}' to '{outputFile}'.");
            }
        }
    }
}
