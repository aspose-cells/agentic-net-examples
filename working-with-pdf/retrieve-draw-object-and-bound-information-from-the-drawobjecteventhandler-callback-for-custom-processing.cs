using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsDrawObjectDemo
{
    // Author: Aspose.Cells .NET example
    // Custom handler to capture draw object details during rendering
    public class CustomDrawObjectEventHandler : DrawObjectEventHandler
    {
        public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
        {
            // Output basic bound information
            Console.WriteLine($"[Draw] Type: {drawObject.Type}, Bounds: X={x}, Y={y}, W={width}, H={height}");
            Console.WriteLine($"       SheetIndex={drawObject.SheetIndex}, Page={drawObject.CurrentPage}/{drawObject.TotalPages}");

            // If the object is associated with a cell, output cell details
            if (drawObject.Cell != null)
            {
                Console.WriteLine($"       Cell: {drawObject.Cell.Name}, Value: {drawObject.Cell.Value}");
            }

            // If the object is a shape, output shape details
            if (drawObject.Shape != null)
            {
                Console.WriteLine($"       Shape: {drawObject.Shape.Name}, Type={drawObject.Shape.Type}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and populate data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Aspose.Cells");
            sheet.Cells["A2"].PutValue("DrawObject Demo");
            sheet.Cells["B1"].Formula = "=A1 & \" \" & A2";

            // Add a rectangle shape
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
            rect.Text = "Sample Shape";

            // ---------- Set up rendering options with the custom handler ----------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true,
                DrawObjectEventHandler = new CustomDrawObjectEventHandler()
            };

            // Render the worksheet to an image (triggers the Draw callback)
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "DrawObjectDemo.png");

            // ---------- Save workbook to PDF using PaginatedSaveOptions ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                DrawObjectEventHandler = new CustomDrawObjectEventHandler()
            };
            workbook.Save("DrawObjectDemo.pdf", pdfOptions);

            // ---------- Save the workbook in native Excel format ----------
            workbook.Save("DrawObjectDemo.xlsx");

            Console.WriteLine("Processing completed.");
        }
    }
}