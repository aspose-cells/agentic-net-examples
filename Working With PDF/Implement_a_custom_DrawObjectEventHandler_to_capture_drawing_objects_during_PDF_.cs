using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectCapture
{
    // Custom handler that records information about each drawn object
    public class CaptureDrawObjectHandler : DrawObjectEventHandler
    {
        // Store captured details for later use or inspection
        public List<string> CapturedInfo { get; } = new List<string>();

        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Build a descriptive string for the current draw object
            string info = $"Type: {drawObject.Type}, Sheet: {drawObject.SheetIndex}, Page: {drawObject.CurrentPage + 1}/{drawObject.TotalPages}, " +
                          $"Bounds: ({x:F2}, {y:F2}, {width:F2}, {height:F2})";

            // Include cell details if the object is a cell
            if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
            {
                info += $", Cell: {drawObject.Cell.Name}, Value: {drawObject.Cell.Value}";
            }

            // Include shape details if the object is a shape
            if (drawObject.Shape != null)
            {
                info += $", Shape: {drawObject.Shape.Name}, Text: {drawObject.Shape.Text}";
            }

            // Store the information
            CapturedInfo.Add(info);

            // Optionally, write to console for immediate feedback
            Console.WriteLine(info);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");
            sheet.Cells["A3"].Formula = "=A1 & \" \" & A2";

            // Add a rectangle shape to generate shape draw objects
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
            rect.Text = "Sample Shape";

            // Create PDF save options and assign the custom draw object handler
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            CaptureDrawObjectHandler handler = new CaptureDrawObjectHandler();
            pdfOptions.DrawObjectEventHandler = handler;

            // Save the workbook to PDF; the handler will be invoked during rendering
            workbook.Save("CapturedDrawObjects.pdf", pdfOptions);

            // After saving, you can inspect the captured information if needed
            Console.WriteLine("\nSummary of captured draw objects:");
            foreach (string entry in handler.CapturedInfo)
            {
                Console.WriteLine(entry);
            }
        }
    }
}