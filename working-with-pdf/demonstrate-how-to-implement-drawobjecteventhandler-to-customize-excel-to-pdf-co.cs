using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectHandlerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with data and a formula
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);
            sheet.Cells["A3"].Formula = "=A2+B2";

            // Add a rectangle shape to the sheet
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 150, 80);
            rect.Text = "Sample Shape";

            // (Optional) Add a picture if you have an image file
            // int picIdx = sheet.Pictures.Add(2, 2, "sample.png");

            // Configure PDF save options and assign a custom DrawObjectEventHandler
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.DrawObjectEventHandler = new CustomDrawObjectHandler();

            // Save the workbook to PDF; the custom handler will be invoked for each draw object
            workbook.Save("CustomDrawObjectDemo.pdf", pdfOptions);
        }

        // Custom handler derived from DrawObjectEventHandler
        private class CustomDrawObjectHandler : DrawObjectEventHandler
        {
            public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
            {
                // Basic information about the drawn object
                Console.WriteLine($"Object Type: {drawObject.Type}");
                Console.WriteLine($"Bounds: X={x}, Y={y}, Width={width}, Height={height}");
                Console.WriteLine($"Sheet Index: {drawObject.SheetIndex}, Page {drawObject.CurrentPage + 1}/{drawObject.TotalPages}");

                // If the object is a cell, output its address and value
                if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
                {
                    Console.WriteLine($"Cell {drawObject.Cell.Name} Value: {drawObject.Cell.Value}");
                }

                // If the object is an image (chart, picture, shape rendered as image), show its byte size
                if (drawObject.Type == DrawObjectEnum.Image && drawObject.ImageBytes != null)
                {
                    Console.WriteLine($"Image byte size: {drawObject.ImageBytes.Length}");
                }

                // If the object is a shape, display its name
                if (drawObject.Shape != null)
                {
                    Console.WriteLine($"Shape Name: {drawObject.Shape.Name}");
                }

                Console.WriteLine(new string('-', 40));
            }
        }
    }
}