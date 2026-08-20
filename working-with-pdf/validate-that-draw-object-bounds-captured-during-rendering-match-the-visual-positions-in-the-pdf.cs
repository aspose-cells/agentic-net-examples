// Title: Validate Shape Bounds During PDF Export with Aspose.Cells for .NET
// Description: A .NET example that adds a rectangle shape to a workbook and uses a custom DrawObjectEventHandler (via PdfSaveOptions) to compare the rendered X, Y, width and height with the shape's original properties, reporting any mismatches within a 0.5‑pixel tolerance.
// Keywords: Aspose.Cells | .NET | C# | PDF export | DrawObjectEventHandler | shape bounds validation | rendering tolerance | Excel to PDF conversion | visual fidelity | custom PdfSaveOptions
// Common Searches: how to verify shape positions when exporting Excel to PDF using Aspose.Cells | Aspose.Cells custom DrawObjectEventHandler example | check rectangle coordinates during PDF generation Aspose.Cells | validate draw object dimensions in PDF output .NET | Aspose.Cells PDF rendering accuracy test
// Developer Intent: Confirm that the coordinates and dimensions of shapes rendered in a PDF match the workbook's shape definitions.
// Use Cases: Automated regression testing to detect layout shifts after Excel‑to‑PDF conversion. | Compliance reporting that requires pixel‑perfect rendering of charts and diagrams. | Debugging complex worksheets by logging shape position discrepancies before publishing PDFs.
// AI Prompts: Create a DrawObjectEventHandler that logs shape bound differences to a JSON file with a configurable tolerance. | Extend the ValidationHandler to also compare cell background colors and borders during PDF export. | Write unit tests for the ValidationHandler that verify no issues for correctly positioned shapes and intentionally offset shapes.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectValidation
{
    // Custom handler that captures draw object bounds and validates them against the original shape properties
    // A .NET example that adds a rectangle shape to a workbook and uses a custom DrawObjectEventHandler (via PdfSaveOptions) to compare the rendered X, Y, width and height with the shape's original properties, reporting any mismatches within a 0.5‑pixel tolerance.
    class ValidationHandler : DrawObjectEventHandler
    {
        // Stores any mismatches found during rendering
        public List<string> Issues { get; } = new List<string>();

        // Tolerance for floating‑point comparison (in pixels)
        private const float Tolerance = 0.5f;

        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Only validate shape draw objects (cells can be validated similarly if needed)
            if (drawObject.Shape != null)
            {
                Shape shape = drawObject.Shape;

                // Shape position and size as defined in the workbook
                float expectedX = shape.Left;
                float expectedY = shape.Top;
                float expectedWidth = shape.Width;
                float expectedHeight = shape.Height;

                // Compare each dimension with a small tolerance
                if (Math.Abs(x - expectedX) > Tolerance ||
                    Math.Abs(y - expectedY) > Tolerance ||
                    Math.Abs(width - expectedWidth) > Tolerance ||
                    Math.Abs(height - expectedHeight) > Tolerance)
                {
                    Issues.Add(
                        $"Shape '{shape.Name}' bounds mismatch. " +
                        $"Expected ({expectedX:F2}, {expectedY:F2}, {expectedWidth:F2}, {expectedHeight:F2}) " +
                        $"but got ({x:F2}, {y:F2}, {width:F2}, {height:F2}).");
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to make the sheet non‑empty
            sheet.Cells["A1"].PutValue("Validation Demo");
            sheet.Cells["A2"].PutValue(12345);

            // Add a rectangle shape whose bounds we will validate
            Shape rect = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                5,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                200, // width (pixels)
                100  // height (pixels)
            );
            rect.Name = "TestRectangle";
            rect.Text = "Validate Me";

            // -------------------- Set up PDF save options with custom handler --------------------
            ValidationHandler handler = new ValidationHandler();

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = handler
            };

            // -------------------- Save workbook to PDF (triggers rendering) --------------------
            workbook.Save("DrawObjectValidation.pdf", pdfOptions);

            // -------------------- Report validation results --------------------
            if (handler.Issues.Count == 0)
            {
                Console.WriteLine("All draw object bounds match the visual positions.");
            }
            else
            {
                Console.WriteLine("Bound mismatches detected:");
                foreach (string issue in handler.Issues)
                {
                    Console.WriteLine(issue);
                }
            }
        }
    }
}
