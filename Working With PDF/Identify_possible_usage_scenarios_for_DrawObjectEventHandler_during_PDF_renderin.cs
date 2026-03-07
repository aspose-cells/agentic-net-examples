using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace DrawObjectEventHandlerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a workbook with sample data and a shape
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Cell A1");
            ws.Cells["B2"].PutValue(123);
            ws.Cells["C3"].Formula = "=A1 & \" \" & B2";

            // Add a rectangle shape to demonstrate shape draw objects
            Shape rect = ws.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 150, 80);
            rect.Text = "Sample Shape";

            // -------------------------------------------------
            // Scenario 1: Log every drawn object (type, position, size)
            // -------------------------------------------------
            PdfSaveOptions logOptions = new PdfSaveOptions();
            logOptions.DrawObjectEventHandler = new LogDrawObjectHandler();
            wb.Save("LogDrawObjects.pdf", logOptions);

            // -------------------------------------------------
            // Scenario 2: Capture image bytes of rendered images (e.g., pictures)
            // -------------------------------------------------
            // Add a tiny picture to the sheet
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
            using (var ms = new System.IO.MemoryStream(pngBytes))
            {
                ws.Pictures.Add(2, 2, ms);
            }

            PdfSaveOptions imageCaptureOptions = new PdfSaveOptions();
            var imageHandler = new ImageBytesCaptureHandler();
            imageCaptureOptions.DrawObjectEventHandler = imageHandler;
            wb.Save("CaptureImageBytes.pdf", imageCaptureOptions);

            // Save the captured image to a separate file (optional)
            if (imageHandler.CapturedImageBytes != null)
            {
                System.IO.File.WriteAllBytes("CapturedImage.png", imageHandler.CapturedImageBytes);
            }

            // -------------------------------------------------
            // Scenario 3: Custom processing for cells (e.g., special handling for A1)
            // -------------------------------------------------
            PdfSaveOptions cellHandlerOptions = new PdfSaveOptions();
            cellHandlerOptions.DrawObjectEventHandler = new CellProcessingHandler();
            wb.Save("CellProcessing.pdf", cellHandlerOptions);
        }

        // Handler that logs information about each draw object during PDF rendering
        private class LogDrawObjectHandler : DrawObjectEventHandler
        {
            public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
            {
                Console.WriteLine($"[Log] Type={drawObject.Type}, Sheet={drawObject.SheetIndex}, Page={drawObject.CurrentPage + 1}/{drawObject.TotalPages}");
                Console.WriteLine($"      Position=({x:F2},{y:F2}), Size=({width:F2}x{height:F2})");
            }
        }

        // Handler that captures image bytes when the draw object represents an image
        private class ImageBytesCaptureHandler : DrawObjectEventHandler
        {
            public byte[] CapturedImageBytes { get; private set; }

            public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
            {
                if (drawObject.Type == DrawObjectEnum.Image && drawObject.ImageBytes != null)
                {
                    CapturedImageBytes = drawObject.ImageBytes;
                    Console.WriteLine($"[ImageCapture] Captured image of {CapturedImageBytes.Length} bytes at ({x:F2},{y:F2})");
                }
            }
        }

        // Handler that performs custom logic for cells, such as detecting a specific cell
        private class CellProcessingHandler : DrawObjectEventHandler
        {
            public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
            {
                if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
                {
                    var cell = drawObject.Cell;
                    Console.WriteLine($"[Cell] Rendering {cell.Name} (Value='{cell.Value}') at ({x:F2},{y:F2})");
                    // Example: special handling for cell A1
                    if (cell.Row == 0 && cell.Column == 0)
                    {
                        Console.WriteLine("   -> This is cell A1 – custom actions could be applied here.");
                    }
                }
            }
        }
    }
}