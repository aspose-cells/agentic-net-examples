// Title: C# – Export Aspose.Cells Worksheet to PDF with Background Colors and Borders
// Description: Shows how to build a workbook, apply a solid background and thin borders to a range, enable ExportDocumentStructure in PdfSaveOptions, calculate formulas, and save the sheet as a PDF that retains all cell formatting.
// Keywords: Aspose.Cells | C# | Export worksheet to PDF | preserve cell formatting | background color | cell borders | PdfSaveOptions | ExportDocumentStructure | style range | PDF conversion example | global
// Common Searches: Aspose.Cells keep cell colors when exporting to PDF | C# export worksheet with borders to PDF | PdfSaveOptions ExportDocumentStructure usage | preserve formatting in PDF using Aspose.Cells | apply style to range before PDF export C#
// Developer Intent: Generate a PDF from a worksheet while maintaining the original cell background colors and border styles.
// Use Cases: Create printable reports or invoices that require exact visual styling in the PDF output. | Archive styled data tables as PDFs for regulatory compliance or record‑keeping. | Automate batch conversion of Excel sheets with custom formatting into shareable PDFs.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to PDF, preserving background colors and borders. | Explain how the ExportDocumentStructure property of PdfSaveOptions influences PDF rendering in Aspose.Cells. | Demonstrate creating a style with solid fill and thin borders, applying it to a range, then saving the workbook as a PDF.

using System;
using System.Drawing;
using Aspose.Cells;

// Shows how to build a workbook, apply a solid background and thin borders to a range, enable ExportDocumentStructure in PdfSaveOptions, calculate formulas, and save the sheet as a PDF that retains all cell formatting.
class ExportWorksheetToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("Item 1");
        worksheet.Cells["B2"].PutValue(123);

        // Create a style with background color and borders
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightYellow;          // background color
        style.Pattern = BackgroundType.Solid;               // apply solid fill

        // Set thin black borders on all sides
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.BottomBorder].Color = Color.Black;
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.LeftBorder].Color = Color.Black;
        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.RightBorder].Color = Color.Black;
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.TopBorder].Color = Color.Black;

        // Apply the style to the desired range (A1:B2)
        StyleFlag flag = new StyleFlag { All = true };
        worksheet.Cells.CreateRange("A1:B2").ApplyStyle(style, flag);

        // Create PDF save options and enable document structure export
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // preserves cell formatting, colors, borders

        // Ensure any formulas are calculated before saving
        workbook.CalculateFormula();

        // Save the worksheet as a PDF file with the specified options
        workbook.Save("WorksheetWithFormatting.pdf", pdfOptions);
    }
}
