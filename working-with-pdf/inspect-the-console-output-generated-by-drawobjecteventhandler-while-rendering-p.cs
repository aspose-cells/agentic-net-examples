using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace DrawObjectEventHandlerDemo
{
    // Custom handler that logs detailed information for each drawn object
    public class LoggingDrawObjectHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Basic object information
            Console.WriteLine($"DrawObject Type: {drawObject.Type}");
            Console.WriteLine($"Position: X={x}, Y={y}, Width={width}, Height={height}");
            Console.WriteLine($"SheetIndex: {drawObject.SheetIndex}, CurrentPage: {drawObject.CurrentPage + 1}/{drawObject.TotalPages}");

            // If the object is a cell, output cell details
            if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
            {
                var cell = drawObject.Cell;
                Console.WriteLine($"Cell Name: {cell.Name}, Value: {cell.Value}, Row: {cell.Row}, Column: {cell.Column}");
            }

            // If the object is an image (e.g., chart, picture, shape), output shape details
            if (drawObject.Type == DrawObjectEnum.Image && drawObject.Shape != null)
            {
                var shape = drawObject.Shape;
                Console.WriteLine($"Shape Name: {shape.Name}, Text: {shape.Text}, Type: {shape.AutoShapeType}");
                Console.WriteLine($"Shape Dimensions: {shape.Width}x{shape.Height} pixels");
            }

            Console.WriteLine(new string('-', 60));
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");
            sheet.Cells["A3"].Formula = "=A1 & \" \" & A2";

            // Add a rectangle shape to generate image draw objects
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
            shape.Text = "Sample Shape";

            // Add a picture to demonstrate image handling
            // (Using a simple 1x1 pixel PNG encoded as base64)
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
            using (var ms = new System.IO.MemoryStream(pngBytes))
            {
                sheet.Pictures.Add(5, 5, ms);
            }

            // Configure PDF save options with the custom draw object handler
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = new LoggingDrawObjectHandler()
            };

            // Save the workbook to PDF; the handler will output details to the console during rendering
            workbook.Save("DrawObjectEventHandlerOutput.pdf", pdfOptions);
        }
    }
}