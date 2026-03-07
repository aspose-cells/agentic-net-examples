using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectDemo
{
    // Custom handler to capture each DrawObject and its bounds during rendering
    public class MyDrawObjectHandler : DrawObjectEventHandler
    {
        // This method is called for every object (cell, image, shape, etc.) rendered
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Output the type of the object and its bounding rectangle
            Console.WriteLine($"DrawObject Type : {drawObject.Type}");
            Console.WriteLine($"Bounds          : X={x}, Y={y}, Width={width}, Height={height}");

            // Additional information based on object type
            if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
            {
                Console.WriteLine($"  Cell Name   : {drawObject.Cell.Name}");
                Console.WriteLine($"  Cell Value  : {drawObject.Cell.Value}");
            }
            else if (drawObject.Type == DrawObjectEnum.Image && drawObject.ImageBytes != null)
            {
                Console.WriteLine($"  Image Size  : {drawObject.ImageBytes.Length} bytes");
            }
            else if (drawObject.Shape != null)
            {
                Console.WriteLine($"  Shape Name  : {drawObject.Shape.Name}");
                Console.WriteLine($"  Shape Text  : {drawObject.Shape.Text}");
            }

            Console.WriteLine(); // Blank line for readability
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells so that cell draw objects are generated
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].Formula = "=A1 & \" \" & B1";

            // Add a shape to generate shape draw objects
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
            shape.Text = "Sample Shape";

            // Configure PDF save options and attach the custom draw object handler
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = new MyDrawObjectHandler()
            };

            // Save the workbook to PDF; during this process the Draw method will be invoked
            workbook.Save("OutputWithDrawObjects.pdf", pdfOptions);
        }
    }
}