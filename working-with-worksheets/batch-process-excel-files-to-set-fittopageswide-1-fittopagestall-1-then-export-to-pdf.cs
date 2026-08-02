// Title: Batch convert Excel (.xlsx) files to single‑page PDF using Aspose.Cells for .NET
// Description: C# code that scans a folder for .xlsx workbooks, sets each worksheet’s PageSetup.FitToPagesWide = 1 and FitToPagesTall = 1, and saves the files as PDFs with PdfSaveOptions.
// Keywords: Aspose.Cells | C# Excel to PDF | FitToPagesWide | FitToPagesTall | batch PDF conversion | page setup | PdfSaveOptions | .NET | single page per sheet | automated Excel export
// Common Searches: Aspose.Cells set FitToPagesWide 1 for all worksheets | C# batch convert xlsx to pdf | export Excel to PDF single page using Aspose | force Excel sheet to fit one page programmatically | save multiple workbooks as PDF C#
// Developer Intent: Process every Excel file in a directory, force each worksheet onto a single page, and generate matching PDF files automatically.
// Use Cases: Create printable PDFs from a collection of financial reports while keeping each sheet on one page. | Provide PDF versions of template workbooks for users without Excel. | Integrate into CI/CD pipelines to produce PDF documentation from Excel data nightly.
// AI Prompts: Generate a reusable C# method that accepts input and output folder paths, applies FitToPagesWide = 1 and FitToPagesTall = 1 to all worksheets, and saves each workbook as PDF with Aspose.Cells. | Add comprehensive error handling and logging to the batch conversion script to capture missing files, load failures, and successful PDF creations. | Show how to customize PdfSaveOptions (e.g., set PDF/A compliance, embed fonts, adjust image quality) while batch converting Excel workbooks to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# code that scans a folder for .xlsx workbooks, sets each worksheet’s PageSetup.FitToPagesWide = 1 and FitToPagesTall = 1, and saves the files as PDFs with PdfSaveOptions.
class BatchExcelToPdf
{
    static void Main()
    {
        // Folder containing source Excel files
        string sourceFolder = @"C:\Input";
        // Folder where PDF files will be saved
        string outputFolder = @"C:\Output";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the source folder
        foreach (string excelPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(excelPath);

            // Set page setup for every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Fit to 1 page wide and 1 page tall (property rules)
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 1;
            }

            // Prepare PDF save options (optional customization)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine PDF output path
            string pdfPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(excelPath) + ".pdf");

            // Save workbook as PDF (save rule)
            workbook.Save(pdfPath, pdfOptions);
        }
    }
}
