// Title: C# – Preserve Excel Cell Borders in PDF with Aspose.Cells RenderSolidGridlines
// Description: Demonstrates how to keep original Excel cell borders when converting a workbook to PDF using Aspose.Cells for .NET. The example creates a styled range, enables gridlines, sets PdfSaveOptions.RenderSolidGridlines to true, and saves the file as a PDF.
// Keywords: Aspose.Cells C# PDF conversion | RenderSolidGridlines true | preserve Excel borders PDF | gridlines PDF output | PdfSaveOptions cell borders | .NET Excel to PDF borders | Aspose.Cells thick borders PDF
// Common Searches: Aspose.Cells keep cell borders when saving as PDF | RenderSolidGridlines property example C# | Excel borders disappear in PDF conversion Aspose | how to show gridlines in PDF with Aspose.Cells | PdfSaveOptions.RenderSolidGridlines usage
// Developer Intent: Set PdfSaveOptions.RenderSolidGridlines = true so the generated PDF retains the workbook's cell borders and gridlines.
// Use Cases: Generating printable invoices where exact border styling must match the Excel source. | Creating archival PDFs of data sheets that require both custom borders and visible gridlines. | Building PDF reports from styled worksheets for distribution without losing visual layout.
// AI Prompts: Provide C# code that uses Aspose.Cells to enable RenderSolidGridlines and export a workbook with thick borders to PDF. | Explain why cell borders are omitted in default Excel‑to‑PDF conversion and how the RenderSolidGridlines option fixes it. | Show a complete Aspose.Cells example that creates a workbook, applies border styles, turns on gridlines, sets RenderSolidGridlines, and saves as PDF.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to keep original Excel cell borders when converting a workbook to PDF using Aspose.Cells for .NET. The example creates a styled range, enables gridlines, sets PdfSaveOptions.RenderSolidGridlines to true, and saves the file as a PDF.
class PreserveCellBordersPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue("Data 2");

            // Define a style with thick black borders
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

            // Apply the border style to the range A1:B2
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B2");
            range.ApplyStyle(borderStyle, new StyleFlag { All = true });

            // Make gridlines visible (optional)
            sheet.IsGridlinesVisible = true;

            // Save the workbook as a PDF file
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save("PreserveBorders.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
