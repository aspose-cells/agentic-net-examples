using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectHandlerDemo
{
    // Custom handler that receives draw events during rendering
    public class CustomDrawObjectHandler : DrawObjectEventHandler
    {
        // This method is called for every draw object (cell, image, shape, etc.)
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Output basic information about the draw object
            Console.WriteLine($"DrawObject Type: {drawObject.Type}");
            Console.WriteLine($"Position: ({x}, {y}) Size: ({width} x {height})");
            Console.WriteLine($"Sheet Index: {drawObject.SheetIndex} Page: {drawObject.CurrentPage + 1}/{drawObject.TotalPages}");

            // If the object is a cell, display its address and value
            if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
            {
                var cell = drawObject.Cell;
                Console.WriteLine($"Cell: {cell.Name} Value: {cell.Value}");
            }

            // If the object is an image (e.g., a picture or chart), indicate its size
            if (drawObject.Type == DrawObjectEnum.Image && drawObject.ImageBytes != null)
            {
                Console.WriteLine($"Image bytes length: {drawObject.ImageBytes.Length}");
            }

            // If the object is a shape, display its name and text
            if (drawObject.Shape != null)
            {
                Console.WriteLine($"Shape Name: {drawObject.Shape.Name}");
                Console.WriteLine($"Shape Text: {drawObject.Shape.Text}");
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

            // Populate some cells with data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].Formula = "=A1 & \" \" & B1";

            // Add a rectangle shape to generate shape draw events
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
            shape.Text = "Sample Shape";

            // Add a picture (if you have an image file, replace the path accordingly)
            // int picIdx = sheet.Pictures.Add(2, 2, "sample.png");

            // Configure PDF save options and attach the custom draw object handler
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = new CustomDrawObjectHandler()
            };

            // Save the workbook to PDF; the Draw method will be invoked during rendering
            workbook.Save("OutputWithDrawObjectHandler.pdf", pdfOptions);

            Console.WriteLine("PDF saved. Check console output for draw object details.");
        }
    }
}