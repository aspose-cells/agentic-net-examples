using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace DrawObjectEventHandlerDemo
{
    // Use case 1: Log every draw object (cells, images, shapes) during PDF rendering
    class LoggingDrawHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            Console.WriteLine($"[Log] Type: {drawObject.Type}, Sheet: {drawObject.SheetIndex}, Page: {drawObject.CurrentPage + 1}/{drawObject.TotalPages}");
            Console.WriteLine($"      Bounds => X:{x}, Y:{y}, W:{width}, H:{height}");

            if (drawObject.Cell != null)
                Console.WriteLine($"      Cell: {drawObject.Cell.Name} = {drawObject.Cell.Value}");

            if (drawObject.Shape != null)
                Console.WriteLine($"      Shape: {drawObject.Shape.Name}, Text: {drawObject.Shape.Text}");
        }
    }

    // Use case 2: Capture image bytes of rendered charts or pictures for further processing
    class ImageCaptureDrawHandler : DrawObjectEventHandler
    {
        public byte[] CapturedImage { get; private set; }

        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            if (drawObject.Type == DrawObjectEnum.Image && drawObject.ImageBytes != null)
            {
                CapturedImage = drawObject.ImageBytes;
                Console.WriteLine($"[ImageCapture] Captured image of size {CapturedImage.Length} bytes at ({x},{y})");
            }
        }
    }

    // Use case 3: Conditionally suppress drawing of a specific cell (e.g., hide cell B2)
    // Note: The handler cannot prevent drawing, but we can replace the cell value temporarily.
    class ConditionalDrawHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
            {
                // Hide content of cell B2 by clearing its value during rendering
                if (drawObject.Cell.Name.Equals("B2", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("[Conditional] Hiding content of cell B2 during PDF rendering.");
                    drawObject.Cell.PutValue(string.Empty);
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook with sample data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue("Sensitive Data"); // This cell will be hidden in use case 3
            sheet.Cells["A3"].PutValue("Data 2");

            // Add a shape
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
            shape.Name = "DemoShape";
            shape.Text = "Sample Shape";

            // Add an image (1x1 red pixel) to demonstrate image capture
            byte[] redPixel = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
            using (MemoryStream imgStream = new MemoryStream(redPixel))
            {
                sheet.Pictures.Add(2, 2, imgStream);
            }

            // ---------- Use case 1: Simple logging ----------
            PdfSaveOptions logOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = new LoggingDrawHandler()
            };
            workbook.Save("LoggingDemo.pdf", logOptions);

            // ---------- Use case 2: Capture image bytes ----------
            ImageCaptureDrawHandler imgHandler = new ImageCaptureDrawHandler();
            PdfSaveOptions imgOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = imgHandler
            };
            workbook.Save("ImageCaptureDemo.pdf", imgOptions);

            // If an image was captured, write it back to the workbook as a new picture
            if (imgHandler.CapturedImage != null)
            {
                using (MemoryStream capturedStream = new MemoryStream(imgHandler.CapturedImage))
                {
                    sheet.Pictures.Add(4, 0, capturedStream);
                }
                workbook.Save("ImageCaptureDemo_WithExtractedImage.pdf", SaveFormat.Xlsx);
            }

            // ---------- Use case 3: Conditional modification ----------
            PdfSaveOptions conditionalOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = new ConditionalDrawHandler()
            };
            workbook.Save("ConditionalDemo.pdf", conditionalOptions);

            Console.WriteLine("PDF rendering completed with all use cases.");
        }
    }
}