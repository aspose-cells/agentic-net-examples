// Title: Capture DrawObject Type and Bounds with a Custom DrawObjectEventHandler in Aspose.Cells .NET
// Description: Demonstrates how to attach a custom DrawObjectEventHandler to ImageOrPrintOptions and PdfSaveOptions, retrieve each DrawObject's type, X/Y coordinates, width, height, sheet index, and page number during PNG or PDF rendering, and log cell or shape details.
// Keywords: Aspose.Cells DrawObjectEventHandler | .NET rendering callback | PDF export draw object bounds | image rendering cell shape info | custom draw object processing | Aspose.Cells worksheet rendering events
// Common Searches: Aspose.Cells get draw object coordinates | How to use DrawObjectEventHandler in C# | Retrieve shape details during PDF export Aspose.Cells | Log cell bounds while rendering worksheet to image | Custom processing of draw objects Aspose.Cells .NET
// Developer Intent: The developer needs to capture the type and bounding rectangle of every draw object (cells, shapes, etc.) during worksheet rendering for custom handling or analysis.
// Use Cases: Audit layout by recording positions and sizes of all cells and shapes in exported PDFs. | Apply dynamic watermarks or overlays based on the exact location of draw objects. | Generate layout statistics or visual guides by analyzing object bounds across pages.
// AI Prompts: Create a DrawObjectEventHandler that writes draw object type, bounds, and cell/shape metadata to a CSV file. | Show how to modify the Draw method to ignore shapes whose names start with "Temp" during rendering. | Provide code that uses the received bounds to draw a custom border around each cell when exporting to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates how to attach a custom DrawObjectEventHandler to ImageOrPrintOptions and PdfSaveOptions, retrieve each DrawObject's type, X/Y coordinates, width, height, sheet index, and page number during PNG or PDF rendering, and log cell or shape details.
class Program
{
    static void Main()
    {
        // Create a new workbook and populate some cells
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Add a shape so that shape draw objects are also generated
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 150, 80);
        shape.Text = "Sample Shape";

        // Set up rendering options with a custom DrawObjectEventHandler
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true,
            DrawObjectEventHandler = new CustomDrawHandler()
        };

        // Render the worksheet to an image – this triggers the Draw method for each draw object
        SheetRender renderer = new SheetRender(sheet, renderOptions);
        renderer.ToImage(0, "RenderedSheet.png");

        // Also demonstrate using the same handler when saving to PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DrawObjectEventHandler = new CustomDrawHandler()
        };
        workbook.Save("RenderedSheet.pdf", pdfOptions);
    }

    // Custom handler that receives the DrawObject and its bounds during rendering
    private class CustomDrawHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Basic information about the draw object
            Console.WriteLine($"DrawObject Type: {drawObject.Type}");
            Console.WriteLine($"Bounds -> X: {x}, Y: {y}, Width: {width}, Height: {height}");
            Console.WriteLine($"SheetIndex: {drawObject.SheetIndex}, Page: {drawObject.CurrentPage + 1}/{drawObject.TotalPages}");

            // If the object represents a cell, output cell details
            if (drawObject.Type == DrawObjectEnum.Cell && drawObject.Cell != null)
            {
                var cell = drawObject.Cell;
                Console.WriteLine($"Cell Name: {cell.Name}, Value: {cell.Value}");
            }

            // If the object represents a shape, output shape details
            if (drawObject.Shape != null)
            {
                var s = drawObject.Shape;
                Console.WriteLine($"Shape Name: {s.Name}, Text: {s.Text}, Width: {s.Width}, Height: {s.Height}");
            }

            Console.WriteLine(); // Separator for readability
        }
    }
}
