using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace DrawObjectEventHandlerDemo
{
    // Custom handler that prints information about each drawn object during PDF rendering
    public class ConsoleDrawObjectHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Basic draw object information
            Console.WriteLine($"DrawObject Type: {drawObject.Type}");
            Console.WriteLine($"Position: X={x}, Y={y}, Width={width}, Height={height}");
            Console.WriteLine($"SheetIndex: {drawObject.SheetIndex}, Page: {drawObject.CurrentPage + 1}/{drawObject.TotalPages}");

            // If the object is a cell, output cell details
            if (drawObject.Cell != null)
            {
                var cell = drawObject.Cell;
                Console.WriteLine($"Cell Name: {cell.Name}, Value: {cell.Value}");
                Console.WriteLine($"Cell Row: {cell.Row}, Column: {cell.Column}");
            }

            // If the object is a shape, output shape details
            if (drawObject.Shape != null)
            {
                var shape = drawObject.Shape;
                Console.WriteLine($"Shape Name: {shape.Name}, Text: {shape.Text}");
                Console.WriteLine($"Shape Type: {shape.AutoShapeType}, Width={shape.Width}, Height={shape.Height}");
            }

            Console.WriteLine(); // Blank line for readability
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells to generate cell draw objects
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");
            sheet.Cells["A3"].Formula = "=A1 & \" \" & A2";

            // Add a shape to generate shape draw objects
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
            shape.Text = "Sample Shape";

            // Configure PDF save options with the custom DrawObjectEventHandler
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = new ConsoleDrawObjectHandler()
            };

            // Save the workbook to PDF; the handler will be invoked during rendering
            workbook.Save("DrawObjectEventHandlerOutput.pdf", pdfOptions);
        }
    }
}