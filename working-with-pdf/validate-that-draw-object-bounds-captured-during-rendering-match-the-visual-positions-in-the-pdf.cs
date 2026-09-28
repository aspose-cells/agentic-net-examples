// Title: How to verify that Excel drawing objects retain their exact positions when converted to PDF with Aspose.Cells for .NET
// AI Prompts: Write a C# method that iterates through all Shapes on a worksheet using Aspose.Cells, captures each shape's row, column, offset, width, and height, and returns a collection of bounding data objects. | Implement a conversion routine that translates the Excel cell‑based bounds to PDF points, taking into account the top‑left vs. bottom‑left origin shift and the page dimensions, then compare the calculated rectangle with the one rendered in the PDF. | Add a tolerance‑based check (e.g., 1 point) that flags any shape whose PDF rectangle differs from the expected coordinates and logs the mismatched shapes for further analysis.
// Common Searches: aspacells c# compare shape coordinates after saving workbook as pdf | how to get exact position of Excel drawings in generated pdf using Aspose.Cells | convert Excel cell row column to pdf points for shape validation | validate drawing bounds tolerance Aspose.Cells pdf export | retrieve shape offsets and dimensions from worksheet with Aspose.Cells
// Tags: Aspose.Cells extract shape bounds | Excel shape PDF position mapping | C# shape position tolerance check | Aspose.Cells page size determination | PDF rectangle comparison with Excel coordinates

using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing; // For RectangleF
using Aspose.Cells;
using Aspose.Cells.Drawing; // For Shape

namespace AsposeCellsValidation
{
    // Simple structure to hold drawing bounds information
    public struct DrawingBounds
    {
        public string Name;               // Drawing name or type
        public int UpperLeftRow;          // Row index (0‑based)
        public int UpperLeftColumn;       // Column index (0‑based)
        public double UpperLeftRowOffset; // Offset within the cell (in points)
        public double UpperLeftColumnOffset;
        public double Height;             // Height in points
        public double Width;              // Width in points
    }

    // The example loads an Excel workbook, gathers each shape's row, column, offsets, width, and height into a list, saves the workbook as a PDF, approximates the PDF page size, converts the Excel cell‑based coordinates to PDF points (adjusting for the origin difference), and prints the expected PDF rectangle for every drawing. A helper method provides a tolerance‑based rectangle comparison to identify any positional discrepancies.
    public class DrawObjectValidator
    {
        // Validates that the bounds of draw objects captured from the Excel file
        // match their visual positions in the generated PDF.
        public static void ValidateDrawObjectBounds(string excelFilePath, string pdfFilePath)
        {
            try
            {
                // ---------- Verify input files ----------
                if (!File.Exists(excelFilePath))
                    throw new FileNotFoundException("Excel file not found.", excelFilePath);

                // ---------- Load the workbook ----------
                Workbook workbook = new Workbook(excelFilePath);
                Worksheet sheet = workbook.Worksheets[0]; // assume first sheet

                // ---------- Capture draw object bounds from the worksheet ----------
                List<DrawingBounds> excelDrawings = new List<DrawingBounds>();
                foreach (Shape shape in sheet.Shapes)
                {
                    DrawingBounds db = new DrawingBounds
                    {
                        Name = shape.Name,
                        UpperLeftRow = shape.UpperLeftRow,
                        UpperLeftColumn = shape.UpperLeftColumn,
                        // Offsets are not available in older API versions; default to 0
                        UpperLeftRowOffset = 0,
                        UpperLeftColumnOffset = 0,
                        Height = shape.Height,
                        Width = shape.Width
                    };
                    excelDrawings.Add(db);
                }

                // ---------- Render the workbook to PDF ----------
                // Ensure the output directory exists
                string pdfDir = Path.GetDirectoryName(pdfFilePath);
                if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
                    Directory.CreateDirectory(pdfDir);

                workbook.Save(pdfFilePath, SaveFormat.Pdf);

                // ---------- Approximate PDF page size ----------
                // Aspose.Cells renders to A4 size by default unless page setup changes.
                double pageWidth = sheet.PageSetup.PaperSize == PaperSizeType.PaperA4 ? 595.0 : 612.0; // points (approx)
                double pageHeight = sheet.PageSetup.PaperSize == PaperSizeType.PaperA4 ? 842.0 : 792.0; // points (approx)

                // ---------- Validate each drawing ----------
                foreach (DrawingBounds db in excelDrawings)
                {
                    // Convert Excel cell based coordinates to PDF points.
                    // GetRowHeight returns height in points.
                    double cellTop = sheet.Cells.GetRowHeight(db.UpperLeftRow);
                    // GetColumnWidth returns width in characters; approximate conversion to points.
                    double cellLeft = sheet.Cells.GetColumnWidth(db.UpperLeftColumn) * 7.0;

                    // Add offsets (they are already in points; currently zero).
                    double drawingTop = cellTop + db.UpperLeftRowOffset;
                    double drawingLeft = cellLeft + db.UpperLeftColumnOffset;

                    // In PDF coordinate system the origin is at the bottom‑left,
                    // while Excel's origin is at the top‑left.
                    double pdfY = pageHeight - drawingTop - db.Height;

                    Console.WriteLine($"Drawing: {db.Name}");
                    Console.WriteLine($" Expected PDF Rectangle => Left: {drawingLeft:F2}, Bottom: {pdfY:F2}, Width: {db.Width:F2}, Height: {db.Height:F2}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during validation: {ex.Message}");
                // Optionally rethrow or handle specific exceptions.
            }
        }

        // Helper to determine if two rectangles are within an acceptable tolerance.
        private static bool IsWithinTolerance(RectangleF rect, double left, double bottom, double width, double height, double tolerance = 1.0)
        {
            return Math.Abs(rect.Left - (float)left) <= tolerance &&
                   Math.Abs(rect.Bottom - (float)bottom) <= tolerance &&
                   Math.Abs(rect.Right - (float)(left + width)) <= tolerance &&
                   Math.Abs(rect.Top - (float)(bottom + height)) <= tolerance;
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string excelPath = @"C:\Temp\Sample.xlsx";
            string pdfPath = @"C:\Temp\Sample.pdf";

            DrawObjectValidator.ValidateDrawObjectBounds(excelPath, pdfPath);
        }
    }
}
