// Title: C# – Preserve Excel Cell Borders in PDF with Aspose.Cells PdfSaveOptions (RenderSolidGridlines)
// Description: Demonstrates how to load or create an Excel workbook, apply thin borders to a range, enable printed gridlines, and configure PdfSaveOptions (GridlineType.Hair, GridlineColor.Black) to render solid gridlines so the generated PDF retains the original cell border formatting.
// Keywords: Aspose.Cells PDF export | PdfSaveOptions RenderSolidGridlines | C# preserve Excel borders PDF | GridlineType Hair Aspose.Cells | Excel to PDF cell borders | .NET Aspose.Cells PDF conversion | PrintGridlines PDF Aspose
// Common Searches: Aspose.Cells keep cell borders when saving as PDF | PdfSaveOptions.RenderSolidGridlines C# example | How to export Excel with solid gridlines using Aspose.Cells | Preserve Excel gridlines in PDF output .NET | Set GridlineType Hair for PDF export Aspose
// Developer Intent: Export an Excel worksheet to PDF while preserving the exact border appearance of the cells.
// Use Cases: Generate PDF reports or invoices where the visual gridlines must match the Excel layout. | Create automated document pipelines that convert styled workbooks to PDF without losing border styling. | Build a web service that receives an XLSX file, applies custom border styles, and returns a PDF with solid black gridlines.
// AI Prompts: Show me a C# snippet that uses Aspose.Cells PdfSaveOptions.RenderSolidGridlines to keep cell borders in the PDF. | Explain how GridlineType.Hair differs from the default gridline setting in Aspose.Cells PDF conversion. | Provide step‑by‑step code to load an Excel file, apply thin borders, enable PrintGridlines, and save as PDF with solid gridlines.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to load or create an Excel workbook, apply thin borders to a range, enable printed gridlines, and configure PdfSaveOptions (GridlineType.Hair, GridlineColor.Black) to render solid gridlines so the generated PDF retains the original cell border formatting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if the file exists)
            Workbook workbook;
            string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Define a style with thin borders on all sides
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the style to the desired range
            StyleFlag flag = new StyleFlag { Borders = true };
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B2");
            range.ApplyStyle(borderStyle, flag);

            // Ensure gridlines are printed (optional, does not affect borders)
            sheet.PageSetup.PrintGridlines = true;

            // Configure PDF save options to render solid gridlines (preserve borders)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                GridlineType = GridlineType.Hair, // solid line instead of default dotted
                GridlineColor = Color.Black       // optional: set gridline color
            };

            // Save the workbook as PDF with the specified options
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
