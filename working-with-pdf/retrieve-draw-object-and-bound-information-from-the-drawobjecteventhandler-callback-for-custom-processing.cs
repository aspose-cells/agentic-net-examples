using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and populate some cells
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Add a shape so that shape draw objects are also generated
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
        shape.Text = "Sample Shape";

        // Create an instance of the custom handler
        var handler = new CaptureDrawObjectHandler();

        // Set rendering options and assign the handler
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true,
            DrawObjectEventHandler = handler
        };

        // Render the worksheet – this triggers the DrawObject events
        SheetRender renderer = new SheetRender(sheet, options);
        renderer.ToImage(0, "RenderedOutput.png");

        // After rendering, output the captured draw object information
        foreach (var info in handler.CapturedObjects)
        {
            Console.WriteLine($"DrawObject Type: {info.DrawObject.Type}");
            Console.WriteLine($"Bounds: X={info.X}, Y={info.Y}, Width={info.Width}, Height={info.Height}");

            // Access cell information if available
            if (info.DrawObject.Cell != null)
            {
                Console.WriteLine($"  Cell: {info.DrawObject.Cell.Name} = {info.DrawObject.Cell.Value}");
            }

            // Access shape information if available
            if (info.DrawObject.Shape != null)
            {
                Console.WriteLine($"  Shape: {info.DrawObject.Shape.Name}, Text=\"{info.DrawObject.Shape.Text}\"");
            }

            Console.WriteLine();
        }
    }

    // Simple container to hold a draw object and its bounds
    class DrawInfo
    {
        public DrawObject DrawObject { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }

    // Custom event handler that captures each draw object during rendering
    class CaptureDrawObjectHandler : DrawObjectEventHandler
    {
        public List<DrawInfo> CapturedObjects { get; } = new List<DrawInfo>();

        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Store the draw object and its bounding rectangle for later processing
            CapturedObjects.Add(new DrawInfo
            {
                DrawObject = drawObject,
                X = x,
                Y = y,
                Width = width,
                Height = height
            });
        }
    }
}