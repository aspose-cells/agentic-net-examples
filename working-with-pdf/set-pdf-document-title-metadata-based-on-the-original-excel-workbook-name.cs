// Title: C# – Set PDF Title from Excel Workbook Name using Aspose.Cells
// Description: Loads an Excel file, extracts its filename (without extension), assigns it to the workbook's built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the workbook as a PDF so the PDF window title matches the original Excel name.
// Keywords: Aspose.Cells PDF title | C# set PDF metadata | DisplayDocTitle option | Excel to PDF title property | built‑in document properties Aspose | save workbook as PDF C# | PDF window title from filename
// Common Searches: set PDF title from Excel file name Aspose.Cells | DisplayDocTitle example C# | assign built‑in Title property before PDF export | extract workbook name without extension C# | Aspose.Cells PDF metadata tutorial
// Developer Intent: Assign the PDF document title so it mirrors the source Excel workbook's filename.
// Use Cases: Create PDF reports where the window title instantly reveals the originating Excel template. | Batch‑convert a directory of .xlsx files to PDFs, automatically embedding each source filename as the PDF title. | Meet document‑management compliance by storing the original workbook name in the PDF title metadata.
// AI Prompts: Show a C# snippet that also sets Author and Subject metadata together with Title using Aspose.Cells. | Provide code to iterate over multiple Excel files, convert each to PDF, and apply DisplayDocTitle so every PDF title matches its source file. | Explain how to programmatically verify that the PDF title metadata was written correctly after saving with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel file, extracts its filename (without extension), assigns it to the workbook's built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the workbook as a PDF so the PDF window title matches the original Excel name.
class SetPdfTitleFromWorkbookName
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Extract the workbook file name without extension
        string title = Path.GetFileNameWithoutExtension(excelPath);

        // Set the built‑in Title property to the extracted name
        workbook.BuiltInDocumentProperties.Title = title;

        // Configure PDF save options to use the document title
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DisplayDocTitle = true   // Ensures PDF window title shows the document title
        };

        // Save the workbook as PDF with the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
