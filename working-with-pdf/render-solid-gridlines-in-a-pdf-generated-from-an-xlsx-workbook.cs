// Title: Render solid hair‑thin gridlines in a PDF from an XLSX workbook using Aspose.Cells for .NET
// Description: Shows how to enable worksheet gridlines, set PdfSaveOptions.GridlineType to Hair and GridlineColor to black, and save the workbook as a PDF with solid thin lines.
// Keywords: Aspose.Cells | PdfSaveOptions | GridlineType | Hair | GridlineColor | solid gridlines | export Excel to PDF | C# | .NET | PDF rendering | worksheet gridlines
// Common Searches: Aspose.Cells export Excel to PDF with solid gridlines | PdfSaveOptions GridlineType Hair example C# | How to set gridline color when saving workbook as PDF | Render thin black gridlines in PDF using Aspose.Cells | C# code to show gridlines in PDF output
// Developer Intent: Generate a PDF from an Excel workbook where the gridlines appear as solid hair‑thin black lines.
// Use Cases: Produce printable financial statements with clear row/column separation. | Create invoices or receipts where thin black gridlines improve table readability. | Export product catalogs or price lists to PDF while preserving consistent gridline styling.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a PDF with solid black hair‑thin gridlines. | Explain the effect of PdfSaveOptions.GridlineType and GridlineColor on PDF gridline appearance. | Show how to enable gridlines on a worksheet and customize their style before converting to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to enable worksheet gridlines, set PdfSaveOptions.GridlineType to Hair and GridlineColor to black, and save the workbook as a PDF with solid thin lines.
class RenderSolidGridlinesPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data so gridlines are visible
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Book");
        sheet.Cells["B2"].PutValue(19.99);
        sheet.Cells["A3"].PutValue("Pen");
        sheet.Cells["B3"].PutValue(2.49);

        // Ensure gridlines are shown in the worksheet
        sheet.IsGridlinesVisible = true;

        // Configure PDF save options to use a solid‑looking (hair) gridline type
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Hair line is a thin solid line; Dotted is the default
            GridlineType = GridlineType.Hair,

            // Optional: set a specific gridline color (e.g., black)
            GridlineColor = System.Drawing.Color.Black
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("SolidGridlines.pdf", pdfOptions);
    }
}
