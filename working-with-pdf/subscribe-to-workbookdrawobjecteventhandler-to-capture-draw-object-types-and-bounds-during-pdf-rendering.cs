using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello World");
        sheet.Cells["B2"].PutValue(12345);

        // Add a shape so that shape draw objects are also generated
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
        shape.Text = "Sample Shape";

        // Configure PDF save options with a custom DrawObjectEventHandler
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            DrawObjectEventHandler = new CaptureDrawObjectHandler()
        };

        // Save the workbook to PDF; the handler will be invoked during rendering
        workbook.Save("CapturedDrawObjects.pdf", saveOptions);
    }

    // Custom handler that captures the type of each draw object and its bounds
    private class CaptureDrawObjectHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Output the draw object type and its position/size
            Console.WriteLine($"DrawObject Type: {drawObject.Type}");
            Console.WriteLine($"Bounds -> X: {x}, Y: {y}, Width: {width}, Height: {height}");
        }
    }
}