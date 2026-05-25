using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using AsposeRange = Aspose.Cells.Range;

class PreserveBordersPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(456);

            // Apply borders to cells to simulate original Excel borders
            Style style = workbook.CreateStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].Color = Color.Black;
            style.Borders[BorderType.BottomBorder].Color = Color.Black;
            style.Borders[BorderType.LeftBorder].Color = Color.Black;
            style.Borders[BorderType.RightBorder].Color = Color.Black;

            // Apply the style to the desired range
            AsposeRange range = sheet.Cells.CreateRange("A1:B3");
            range.ApplyStyle(style, new StyleFlag { All = true });

            // Ensure gridlines are visible (optional)
            sheet.IsGridlinesVisible = true;
            sheet.PageSetup.PrintGridlines = true;

            // Configure PDF save options to render solid gridlines (preserve original borders)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                GridlineType = GridlineType.Hair, // solid‑like gridlines
                GridlineColor = Color.Black        // preserve original border color
            };

            // Define output file path
            string outputPath = "PreserveBorders.pdf";

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}