// Title: Capture Draw Object Types and Bounds with Aspose.Cells DrawObjectEventHandler during PDF Export (C#)
// Description: This example creates a workbook, adds data and a rectangle shape, then sets PdfSaveOptions with a custom DrawObjectEventHandler. When the workbook is saved as PDF, the overridden Draw method logs each draw object's type and its X, Y, width, and height values, enabling developers to inspect rendering details.
// Keywords: Aspose.Cells | DrawObjectEventHandler | PDF export | C# | .NET | capture draw object bounds | log draw object type | custom PdfSaveOptions | shape rendering | PDF rendering events
// Common Searches: Aspose.Cells capture draw object coordinates PDF | DrawObjectEventHandler example C# | log shape bounds during PDF save Aspose.Cells | custom PDF rendering events Aspose.Cells .NET
// Developer Intent: Subscribe to DrawObjectEventHandler to record the type and bounding rectangle of each object rendered while converting a workbook to PDF.
// Use Cases: Debug layout issues by printing draw object positions in the generated PDF. | Gather metrics on shapes and images rendered for analytics or reporting. | Apply custom transformations or replacements based on object type and bounds during PDF conversion.
// AI Prompts: Show how to store draw object information in a list instead of writing to the console. | Provide code that filters only shape draw objects in the Draw method and skips other types. | Explain how to access additional properties of DrawObject, such as image data, within the event handler.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This example creates a workbook, adds data and a rectangle shape, then sets PdfSaveOptions with a custom DrawObjectEventHandler. When the workbook is saved as PDF, the overridden Draw method logs each draw object's type and its X, Y, width, and height values, enabling developers to inspect rendering details.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["A2"].PutValue("World");

        // Add a shape to generate shape‑type draw objects
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
        shape.Text = "Sample Shape";

        // Configure PDF save options with a custom DrawObjectEventHandler
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DrawObjectEventHandler = new CaptureDrawObjectHandler()
        };

        // Save the workbook to PDF; the handler will be invoked during rendering
        workbook.Save("CapturedDrawObjects.pdf", pdfOptions);
    }

    // Custom handler that logs the type and bounds of each drawn object
    private class CaptureDrawObjectHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            Console.WriteLine($"DrawObject Type: {drawObject.Type}");
            Console.WriteLine($"Bounds -> X: {x}, Y: {y}, Width: {width}, Height: {height}");
        }
    }
}
