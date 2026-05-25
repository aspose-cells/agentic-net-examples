using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectValidation
{
    // Holds captured draw object information during rendering
    class CapturedInfo
    {
        public DrawObjectEnum Type { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public Shape Shape { get; set; }
        public Cell Cell { get; set; }
    }

    // Custom handler that records bounds of each draw object
    class ValidationDrawObjectHandler : DrawObjectEventHandler
    {
        public List<CapturedInfo> CapturedObjects { get; } = new List<CapturedInfo>();

        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Store the information for later validation
            CapturedObjects.Add(new CapturedInfo
            {
                Type = drawObject.Type,
                X = x,
                Y = y,
                Width = width,
                Height = height,
                Shape = drawObject.Shape,
                Cell = drawObject.Cell
            });
        }
    }

    class Program
    {
        static void Main()
        {
            // ------------------- Create workbook and content -------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B2"].PutValue(12345);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Add a rectangle shape (will be rendered as an image draw object)
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 150, 80);
            rect.Text = "Sample Shape";

            // ------------------- Set up rendering with custom handler -------------------
            ValidationDrawObjectHandler handler = new ValidationDrawObjectHandler();

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = handler
            };

            // Render workbook to PDF (this triggers the Draw method of the handler)
            workbook.Save("RenderedDocument.pdf", pdfOptions);

            // ------------------- Validate captured bounds -------------------
            const float tolerance = 0.5f; // acceptable difference in points

            foreach (var info in handler.CapturedObjects)
            {
                // Validation for shape objects (rendered as Image type)
                if (info.Type == DrawObjectEnum.Image && info.Shape != null)
                {
                    // Shape's own coordinates are in pixels; Aspose.Cells converts them to points during rendering.
                    // For a simple comparison we use the shape's Left/Top/Width/Height properties.
                    float shapeLeft = info.Shape.Left;
                    float shapeTop = info.Shape.Top;
                    float shapeWidth = info.Shape.Width;
                    float shapeHeight = info.Shape.Height;

                    bool leftMatch = Math.Abs(info.X - shapeLeft) <= tolerance;
                    bool topMatch = Math.Abs(info.Y - shapeTop) <= tolerance;
                    bool widthMatch = Math.Abs(info.Width - shapeWidth) <= tolerance;
                    bool heightMatch = Math.Abs(info.Height - shapeHeight) <= tolerance;

                    Console.WriteLine($"Shape \"{info.Shape.Name}\" validation:");
                    Console.WriteLine($"  Position X match: {leftMatch} (captured={info.X}, shape={shapeLeft})");
                    Console.WriteLine($"  Position Y match: {topMatch} (captured={info.Y}, shape={shapeTop})");
                    Console.WriteLine($"  Width match:      {widthMatch} (captured={info.Width}, shape={shapeWidth})");
                    Console.WriteLine($"  Height match:     {heightMatch} (captured={info.Height}, shape={shapeHeight})");
                }

                // Validation for cell objects
                if (info.Type == DrawObjectEnum.Cell && info.Cell != null)
                {
                    // Cells are rendered based on their row/column indices.
                    // We can approximate the expected position using the cell's row height and column width.
                    // For demonstration, we simply output the captured bounds.
                    Console.WriteLine($"Cell \"{info.Cell.Name}\" rendered at ({info.X}, {info.Y}) size {info.Width}x{info.Height}");
                }
            }

            // ------------------- Save workbook (optional) -------------------
            workbook.Save("ValidatedWorkbook.xlsx");
        }
    }
}