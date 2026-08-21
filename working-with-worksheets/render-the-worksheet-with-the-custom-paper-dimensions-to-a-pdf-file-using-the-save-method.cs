// Title: Export a Worksheet with a 5×7‑inch Custom Paper Size to PDF using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates cells A1‑B3, applies a 5 in × 7 in custom paper size via PageSetup.CustomPaperSize, and saves the workbook as a PDF with Workbook.Save (SaveFormat.Pdf).
// Keywords: Aspose.Cells | C# | .NET | custom paper size | PageSetup.CustomPaperSize | PDF export | SaveFormat.Pdf | 5x7 inch | worksheet to PDF | Excel to PDF rendering
// Common Searches: Aspose.Cells set custom paper size and export to PDF | C# export Excel worksheet as 5x7 inch PDF | PageSetup.CustomPaperSize example Aspose.Cells | Save workbook as PDF with specific dimensions .NET | Render worksheet to PDF with non‑standard page size
// Developer Intent: Export a worksheet to a PDF file using a 5 × 7‑inch custom page size.
// Use Cases: Generate 5×7‑inch product flyers directly from Excel data. | Create compact invoices or receipts that match a specific envelope size. | Produce mobile‑friendly PDF reports with a reduced page footprint.
// AI Prompts: Show how to set a 5 in × 7 in custom paper size for a worksheet and save it as PDF with Aspose.Cells in C#. | Provide a C# example that configures margins, orientation, and scaling together with a custom paper size before PDF export. | Explain how to use rendering options to ensure worksheet content fits a 5×7‑inch PDF page when using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, populates cells A1‑B3, applies a 5 in × 7 in custom paper size via PageSetup.CustomPaperSize, and saves the workbook as a PDF with Workbook.Save (SaveFormat.Pdf).
class RenderWorksheetCustomPaperToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apples");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Bananas");
        worksheet.Cells["B3"].PutValue(20);

        // Set custom paper size (width: 5 inches, height: 7 inches)
        worksheet.PageSetup.CustomPaperSize(5.0, 7.0);

        // Save the worksheet to a PDF file using the Save method
        workbook.Save("CustomPaperWorksheet.pdf", SaveFormat.Pdf);
    }
}
