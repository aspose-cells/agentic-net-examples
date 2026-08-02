// Title: Render Solid Gridlines in PDF from Excel using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, enable gridlines, set PrintGridlines, configure PdfSaveOptions (GridlineType.Hair, GridlineColor), and save the file as a PDF that displays solid black gridlines.
// Keywords: Aspose.Cells PDF export | C# gridlines PDF | .NET Excel to PDF | PrintGridlines Aspose.Cells | GridlineType Hair | solid gridlines PDF | gridline color Aspose.Cells | Excel workbook PDF conversion | cell borders PDF | Aspose.Cells PDFSaveOptions
// Common Searches: Aspose.Cells how to show gridlines in PDF | C# export Excel to PDF with solid lines | PrintGridlines true PDF output Aspose | Set gridline thickness and color when saving to PDF | Convert XLSX to PDF with visible borders .NET
// Developer Intent: Export an Excel worksheet to PDF with visible solid gridlines using Aspose.Cells.
// Use Cases: Generate printable reports where cell borders must be clearly visible. | Apply corporate styling by customizing gridline type and color in PDF exports. | Create invoices or price lists in PDF where gridlines separate line items for readability.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet to PDF showing thick black gridlines. | Explain the impact of PrintGridlines and GridlineType settings on PDF rendering in Aspose.Cells. | Show how to change gridline color to blue and use a dotted style when saving a workbook as PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfGridlinesDemo
{
    // Shows how to create a workbook, enable gridlines, set PrintGridlines, configure PdfSaveOptions (GridlineType.Hair, GridlineColor), and save the file as a PDF that displays solid black gridlines.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data so that gridlines are visible
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);

            // Ensure gridlines are shown on the sheet
            sheet.IsGridlinesVisible = true;

            // Instruct the printer layout to actually print the gridlines.
            // When PDF is generated this results in solid gridlines.
            sheet.PageSetup.PrintGridlines = true;

            // Create PDF save options.
            // Set GridlineType to Hair (thin line) – this is the closest to a solid line.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                GridlineType = GridlineType.Hair,
                // Optional: change gridline color if desired
                GridlineColor = Color.Black
            };

            // Save the workbook as PDF with the specified options.
            workbook.Save("WorkbookWithSolidGridlines.pdf", pdfOptions);
        }
    }
}
