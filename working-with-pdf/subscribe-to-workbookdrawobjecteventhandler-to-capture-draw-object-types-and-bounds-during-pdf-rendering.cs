// Title: Log Draw Object Types and Bounds During PDF Export with Aspose.Cells for .NET
// Description: Demonstrates how to attach a custom DrawObjectEventHandler to PdfSaveOptions so that each visual element (cells, charts, shapes, images) rendered during workbook‑to‑PDF conversion is logged with its type and bounding rectangle. The example creates a workbook, adds data and a rectangle shape, configures the handler, and saves the file as PDF, outputting object details to the console.
// Keywords: Aspose.Cells | DrawObjectEventHandler | PDF rendering | C# PDF export | PdfSaveOptions | capture draw object bounds | log draw objects | Aspose.Cells .NET | shape bounds PDF | event handler Aspose
// Common Searches: Aspose.Cells capture draw object coordinates PDF conversion | C# DrawObjectEventHandler example | Log PDF rendering objects Aspose.Cells | Get shape bounds when saving workbook to PDF | Aspose.Cells PDF save options draw object event
// Developer Intent: Attach a custom DrawObjectEventHandler to PdfSaveOptions to record each object's type and bounding rectangle while converting a workbook to PDF.
// Use Cases: Diagnose layout mismatches by reviewing exact positions of cells, charts, and shapes in the generated PDF. | Generate a compliance or audit report of all visual elements exported to PDF. | Collect draw‑object metrics for analytics on element usage across workbooks. | Integrate draw‑object logging with monitoring tools such as Serilog or NLog in production PDF pipelines.
// AI Prompts: Write a C# snippet that saves the draw‑object log to a CSV file using Aspose.Cells. | Show how to modify CustomDrawObjectHandler to capture only images and ignore text. | Explain how to combine DrawObjectEventHandler with a third‑party logging framework like Serilog. | Provide a PowerShell script to run the example and capture console output. | Create a unit test that verifies the handler receives at least one shape draw object.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsDrawObjectDemo
{
    // Custom handler that captures the type of each drawn object and its bounds
    // Demonstrates how to attach a custom DrawObjectEventHandler to PdfSaveOptions so that each visual element (cells, charts, shapes, images) rendered during workbook‑to‑PDF conversion is logged with its type and bounding rectangle. The example creates a workbook, adds data and a rectangle shape, configures the handler, and saves the file as PDF, outputting object details to the console.
    public class CustomDrawObjectHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Log the object type and its bounding rectangle
            Console.WriteLine($"DrawObject Type: {drawObject.Type}");
            Console.WriteLine($"Bounds -> X: {x}, Y: {y}, Width: {width}, Height: {height}");
            Console.WriteLine();
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and populate it with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello Aspose.Cells");
                sheet.Cells["B2"].PutValue(12345);
                sheet.Cells["C3"].Formula = "=A1 & \" - \" & B2";

                // Add a shape to generate shape draw objects as well
                var shape = sheet.Shapes.AddShape(
                    Aspose.Cells.Drawing.MsoDrawingType.Rectangle, // shape type
                    5,   // upper left row
                    0,   // upper left column
                    5,   // top offset in pixels
                    0,   // left offset in pixels
                    150, // width in pixels
                    80   // height in pixels
                );
                shape.Text = "Sample Shape";

                // Configure PDF save options and attach the custom draw object handler
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    DrawObjectEventHandler = new CustomDrawObjectHandler()
                };

                // Save the workbook to PDF; the handler will be invoked during rendering
                workbook.Save("DrawObjectCaptureDemo.pdf", pdfOptions);
                Console.WriteLine("PDF saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
