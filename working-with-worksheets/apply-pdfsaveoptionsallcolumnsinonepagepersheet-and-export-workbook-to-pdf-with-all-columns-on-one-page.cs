// Title: C# – Export Aspose.Cells Workbook to PDF with All Columns on One Page per Sheet
// Description: This example builds a workbook, populates 50 columns with sample data, and applies PdfSaveOptions with AllColumnsInOnePagePerSheet (and optionally OnePagePerSheet) set to true, generating a PDF where each worksheet’s columns are compressed to fit a single page.
// Keywords: Aspose.Cells PDF export | AllColumnsInOnePagePerSheet | OnePagePerSheet | C# Aspose.Cells example | fit columns PDF | scale worksheet to one page | PdfSaveOptions C#
// Common Searches: Aspose.Cells fit all columns on one PDF page | PdfSaveOptions AllColumnsInOnePagePerSheet C# | Export large worksheet to single-page PDF Aspose | C# code to set OnePagePerSheet Aspose.Cells | How to scale worksheet columns when saving as PDF
// Developer Intent: Create a PDF where every sheet’s columns are compressed to a single page.
// Use Cases: Printing wide data tables without line breaks | Generating compact PDF reports from spreadsheets | Delivering invoices where all columns must appear on one page | Archiving dashboards with many metrics in a single-page PDF | Sharing data extracts that need to stay on one printable page
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to PDF with AllColumnsInOnePagePerSheet enabled. | Show how to combine AllColumnsInOnePagePerSheet and OnePagePerSheet in PdfSaveOptions. | Explain the difference between AllColumnsInOnePagePerSheet and OnePagePerSheet when exporting to PDF. | Provide a step-by-step guide to fit all worksheet columns onto one PDF page using Aspose.Cells. | Generate a minimal example that demonstrates scaling columns to one page in a PDF output.

using System;
using Aspose.Cells;

// This example builds a workbook, populates 50 columns with sample data, and applies PdfSaveOptions with AllColumnsInOnePagePerSheet (and optionally OnePagePerSheet) set to true, generating a PDF where each worksheet’s columns are compressed to fit a single page.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data across many columns to demonstrate column fitting
        for (int col = 0; col < 50; col++)
        {
            sheet.Cells[0, col].PutValue("Header " + (col + 1));
            sheet.Cells[1, col].PutValue("Data " + (col + 1));
        }

        // Set PDF save options to place all columns of each sheet on a single page
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.AllColumnsInOnePagePerSheet = true;
        // Optional: ensure the entire sheet fits on one page (both columns and rows)
        pdfOptions.OnePagePerSheet = true;

        // Save the workbook as a PDF file using the configured options
        workbook.Save("AllColumnsOnePage.pdf", pdfOptions);
    }
}
