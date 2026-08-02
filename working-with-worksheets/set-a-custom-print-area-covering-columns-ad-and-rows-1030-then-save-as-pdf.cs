// Title: Set Print Area A10:D30 and Export Worksheet to PDF with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates rows 10‑30 and columns A‑D, assigns the range A10:D30 as the print area via PageSetup.PrintArea, and saves the sheet as a PDF using PdfSaveOptions so only the defined region is rendered.
// Keywords: Aspose.Cells | C# | PrintArea | A10:D30 | PDF export | PageSetup | PdfSaveOptions | custom print range | worksheet to PDF
// Common Searches: Aspose.Cells set print area C# | Export specific range to PDF Aspose.Cells | PageSetup.PrintArea example | Save worksheet as PDF with defined print area | C# Aspose.Cells PDF only selected cells
// Developer Intent: Define a worksheet print region and generate a PDF that contains only that region using Aspose.Cells.
// Use Cases: Produce a PDF report that includes only the data table located in rows 10‑30, columns A‑D. | Create printable invoices or statements by limiting the output to a predefined range before exporting. | Generate a preview PDF of a subset of a large sheet for documentation or review purposes.
// AI Prompts: Write C# code with Aspose.Cells to set the print area to B5:F20, change orientation to landscape, and save as PDF. | Show how to assign different print areas to multiple worksheets and export each to its own PDF file using Aspose.Cells. | Explain how to adjust page margins and scaling when exporting a custom print area to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Creates a workbook, populates rows 10‑30 and columns A‑D, assigns the range A10:D30 as the print area via PageSetup.PrintArea, and saves the sheet as a PDF using PdfSaveOptions so only the defined region is rendered.
class SetPrintAreaAndSavePdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Populate some data within the print area for demonstration
        for (int row = 10; row <= 30; row++)
        {
            for (int col = 0; col < 4; col++) // Columns A (0) to D (3)
            {
                worksheet.Cells[row - 1, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Set the custom print area to columns A‑D and rows 10‑30
        worksheet.PageSetup.PrintArea = "A10:D30";

        // Save the workbook as PDF; the defined print area will be used during printing
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("CustomPrintArea.pdf", pdfOptions);
    }
}
