// Title: How to Preserve Cell Background Colors When Converting Excel to PDF with Aspose.Cells for .NET
// Description: Demonstrates creating a workbook, applying a solid fill style to a cell, and exporting it to PDF using Aspose.Cells. The default PdfSaveOptions keep the cell's background shading intact in the generated PDF.
// Keywords: Aspose.Cells PDF export | C# preserve cell color | Excel to PDF background shading | PdfSaveOptions Aspose | solid fill style Aspose.Cells
// Common Searches: Aspose.Cells keep cell background when saving to PDF | C# export Excel with colors to PDF | PdfSaveOptions retain shading Aspose | how to preserve cell fill in PDF using Aspose.Cells | Excel cell background color lost in PDF conversion
// Developer Intent: Export an Excel workbook to PDF while retaining the visual background colors of cells.
// Use Cases: Generate a status report where colored cells convey meaning and need to appear unchanged in the PDF. | Create printable invoices with highlighted rows that must keep their background hues after conversion. | Automate batch conversion of spreadsheets containing conditional formatting colors to PDF without losing visual cues.
// AI Prompts: Write C# code that sets a solid background color for a cell using Aspose.Cells and saves the workbook to PDF with the color preserved. | Explain which PdfSaveOptions settings affect background shading retention in Aspose.Cells for .NET. | Show an example of exporting a worksheet with multiple colored cells to PDF and verify that all fills are retained.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

// Demonstrates creating a workbook, applying a solid fill style to a cell, and exporting it to PDF using Aspose.Cells. The default PdfSaveOptions keep the cell's background shading intact in the generated PDF.
class PreserveCellBackgroundInPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a value into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Background Color Demo");

            // Create a style with solid fill and set the foreground color (the visible background)
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.Yellow; // desired cell background color

            // Apply the style to the cell
            cell.SetStyle(style);

            // Set PDF save options (default behavior preserves cell shading)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            string outputPath = "BackgroundPreserved.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
