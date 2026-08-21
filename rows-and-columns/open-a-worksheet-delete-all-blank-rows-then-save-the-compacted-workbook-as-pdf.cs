// Title: C# – Delete Blank Rows in an Excel Worksheet and Export to PDF with Aspose.Cells
// Description: Load an Excel file using Aspose.Cells, remove every empty row with Cells.DeleteBlankRows, configure PdfSaveOptions to ignore blank pages, and save the cleaned workbook as a PDF.
// Keywords: Aspose.Cells | C# delete blank rows | Excel to PDF conversion | PdfSaveOptions | IgnoreBlank pages | Remove empty rows | Cells.DeleteBlankRows | Workbook.Save PDF | Aspose.Cells API
// Common Searches: Aspose.Cells delete blank rows C# | How to remove empty rows before PDF export Aspose | Convert Excel to PDF ignoring blank pages .NET | Cells.DeleteBlankRows example | PdfSaveOptions PrintingPageType.IgnoreBlank usage
// Developer Intent: Clean a worksheet by deleting all empty rows and generate a compact PDF without blank pages using Aspose.Cells for .NET.
// Use Cases: Prepare data‑driven reports by stripping out blank rows prior to PDF generation, ensuring a tidy printable document. | Archive Excel workbooks as PDFs after removing unnecessary empty rows to reduce file size and improve readability. | Create client‑facing PDFs from spreadsheets that contain sporadic gaps, preventing the appearance of blank pages in the final output.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, deletes all blank rows, sets PdfSaveOptions to ignore blank pages, and saves the result as a PDF. | Show an example of using Cells.DeleteBlankRows and PdfSaveOptions.PrintingPageType = IgnoreBlank to convert an Excel worksheet to a compact PDF. | Explain how to programmatically remove empty rows from a specific worksheet before exporting to PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an Excel file using Aspose.Cells, remove every empty row with Cells.DeleteBlankRows, configure PdfSaveOptions to ignore blank pages, and save the cleaned workbook as a PDF.
class DeleteBlankRowsAndSavePdf
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (you can modify this to target a specific sheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete all blank rows in the worksheet
        worksheet.Cells.DeleteBlankRows();

        // Optional: configure PDF save options (e.g., ignore completely blank pages)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PrintingPageType = PrintingPageType.IgnoreBlank
        };

        // Save the compacted workbook as a PDF file
        workbook.Save(outputPath, pdfOptions);
    }
}
