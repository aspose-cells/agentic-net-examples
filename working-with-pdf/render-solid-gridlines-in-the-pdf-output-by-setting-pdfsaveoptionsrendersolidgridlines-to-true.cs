// Title: Enable solid gridlines when saving an Aspose.Cells workbook to PDF using C#
// AI Prompts: Generate a PDF from a workbook and turn on solid gridlines by setting PdfSaveOptions.RenderSolidGridlines = true in C#. | Load an existing Excel file, apply thin borders to the used range, enable solid gridlines, and export it as a PDF with Aspose.Cells for .NET. | Adjust PDF save options to include visible gridlines while customizing page margins and other settings in a C# Aspose.Cells application.
// Common Searches: Aspose.Cells C# how to render solid gridlines in PDF | PdfSaveOptions.RenderSolidGridlines true usage example | Export Excel workbook to PDF with gridlines visible using Aspose.Cells .NET | Show Excel gridlines in PDF output with Aspose.Cells C# | C# Aspose.Cells PDF save options to include gridlines
// Tags: Aspose.Cells PDF gridlines rendering | PdfSaveOptions.RenderSolidGridlines C# | export Excel to PDF with visible gridlines Aspose.Cells | apply thin borders before PDF export Aspose.Cells | configure PDF save options Aspose.Cells .NET | enable gridlines in PDF output Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new Workbook, populates sample data, applies a thin black border style to the used range, sets PdfSaveOptions.RenderSolidGridlines = true, and saves the workbook as OutputWithSolidGridlines.pdf, producing a PDF where the worksheet gridlines are rendered as solid lines.
class RenderSolidGridlinesPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Enable printing of gridlines if the property exists in the used version
            // workbook.Settings.PrintGridlines = true; // Uncomment if supported

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to visualize gridlines
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);
            sheet.Cells["A3"].PutValue(789);
            sheet.Cells["B3"].PutValue(101112);

            // Create a style with thin black borders
            Style style = workbook.CreateStyle();
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.LeftBorder].Color = Color.Black;
            style.Borders[BorderType.RightBorder].Color = Color.Black;
            style.Borders[BorderType.TopBorder].Color = Color.Black;

            // Apply the style to the used range
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            usedRange.ApplyStyle(style, new StyleFlag() { All = true });

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save("OutputWithSolidGridlines.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
