// Title: Export Each Excel Worksheet to Separate PDF with OnePagePerSheet – Aspose.Cells C#
// Description: Learn how to load an Excel workbook with Aspose.Cells, enable PdfSaveOptions.OnePagePerSheet, iterate through worksheets, set SheetSet for each sheet, and save every worksheet as an individual PDF file named after the sheet.
// Keywords: Aspose.Cells PDF export C# | OnePagePerSheet option | save worksheet as PDF | SheetSet per sheet PDF | C# Excel to PDF batch conversion | individual PDF per worksheet
// Common Searches: Aspose.Cells export each sheet to separate PDF | C# OnePagePerSheet true PDF conversion | How to use SheetSet with PdfSaveOptions | Save Excel worksheets as individual PDFs .NET | Batch convert Excel workbook to multiple PDFs
// Developer Intent: Create a separate PDF file for every worksheet in an Excel workbook, ensuring each PDF contains only one page.
// Use Cases: Generate distinct PDF reports for each worksheet in a multi‑sheet workbook. | Automate archival of individual sheet PDFs for compliance or record‑keeping. | Prepare per‑sheet PDFs for printing, guaranteeing a single‑page layout per document.
// AI Prompts: Provide C# code using Aspose.Cells that sets PdfSaveOptions.OnePagePerSheet to true and exports each worksheet to its own PDF file. | Show how to configure SheetSet in PdfSaveOptions to limit the save operation to a specific worksheet. | Explain how to sanitize worksheet names for valid file names when saving each sheet as a PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Learn how to load an Excel workbook with Aspose.Cells, enable PdfSaveOptions.OnePagePerSheet, iterate through worksheets, set SheetSet for each sheet, and save every worksheet as an individual PDF file named after the sheet.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options and enable OnePagePerSheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true;

        // Iterate through all worksheets
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            // Restrict the save operation to the current worksheet
            pdfOptions.SheetSet = new SheetSet(i);

            // Generate a file name based on the worksheet name
            string sheetName = workbook.Worksheets[i].Name;
            string outputFile = $"{sheetName}.pdf";

            // Save the current worksheet as an individual PDF file
            workbook.Save(outputFile, pdfOptions);
        }
    }
}
