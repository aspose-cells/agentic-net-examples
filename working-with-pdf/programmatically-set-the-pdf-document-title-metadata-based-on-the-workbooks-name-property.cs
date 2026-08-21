// Title: Set PDF Title from Excel Workbook Name with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, extracts the file name (without extension), assigns it to the built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the workbook as a PDF so viewers display the title metadata derived from the source workbook.
// Keywords: Aspose.Cells | C# | PDF title metadata | DisplayDocTitle | BuiltInDocumentProperties.Title | Excel to PDF conversion | set PDF document title | programmatic PDF metadata | save PDF with title | Aspose.Cells PDF options
// Common Searches: Aspose.Cells set PDF title from workbook name | C# display document title in PDF using Aspose.Cells | PdfSaveOptions DisplayDocTitle example | How to set BuiltInDocumentProperties.Title before saving PDF | Set PDF metadata programmatically Aspose.Cells
// Developer Intent: Assign the workbook filename as the PDF document title before saving.
// Use Cases: Generate PDFs whose title matches the source Excel filename for improved cataloging. | Batch convert a folder of .xlsx files to PDFs with automatic title metadata. | Ensure PDF viewers display a custom title by enabling DisplayDocTitle. | Integrate title setting into automated reporting pipelines.
// AI Prompts: Generate C# code that sets author, subject, and title metadata for a PDF using Aspose.Cells. | Write a script to batch process a directory of .xlsx files, converting each to PDF and setting the PDF title to the original filename. | Explain how PdfSaveOptions.DisplayDocTitle affects PDF viewers and how to verify the title property after export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, extracts the file name (without extension), assigns it to the built‑in Title property, enables DisplayDocTitle in PdfSaveOptions, and saves the workbook as a PDF so viewers display the title metadata derived from the source workbook.
class PdfTitleFromWorkbookName
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(excelPath);

        // Derive the workbook name from the file name (without extension)
        string workbookName = Path.GetFileNameWithoutExtension(excelPath);

        // Set the built‑in document title to the workbook name
        workbook.BuiltInDocumentProperties.Title = workbookName;

        // Create PDF save options and enable displaying the document title
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.DisplayDocTitle = true;   // ensures the PDF viewer shows the title

        // Save the workbook as PDF (save rule)
        workbook.Save("output.pdf", pdfOptions);
    }
}
