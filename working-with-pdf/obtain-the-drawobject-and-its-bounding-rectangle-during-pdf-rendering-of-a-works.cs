using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and populate some cells
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["A2"].PutValue("World");
        sheet.Cells["B1"].PutValue(12345);

        // Add a shape so that shape draw objects are also generated
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 120, 60);
        shape.Text = "Sample Shape";

        // Prepare PDF save options and attach a custom DrawObjectEventHandler
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        var collector = new DrawObjectCollector();
        pdfOptions.DrawObjectEventHandler = collector;

        // Save the workbook to PDF – this triggers the Draw method for each rendered object
        workbook.Save("RenderedOutput.pdf", pdfOptions);

        // After rendering, output the collected draw objects and their bounding rectangles
        foreach (var info in collector.Collected)
        {
            Console.WriteLine($"Object Type: {info.DrawObject.Type}");
            Console.WriteLine($"Sheet Index: {info.DrawObject.SheetIndex}, Page: {info.DrawObject.CurrentPage + 1}/{info.DrawObject.TotalPages}");
            Console.WriteLine($"Bounds -> X: {info.X}, Y: {info.Y}, Width: {info.Width}, Height: {info.Height}");
            Console.WriteLine(new string('-', 40));
        }
    }
}

// Simple container to hold a DrawObject together with its bounding rectangle
public class DrawObjectInfo
{
    public DrawObject DrawObject { get; set; }
    public float X { get; set; }
    public float Y { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
}

// Custom handler that captures each DrawObject and its rectangle during rendering
public class DrawObjectCollector : DrawObjectEventHandler
{
    public List<DrawObjectInfo> Collected { get; } = new List<DrawObjectInfo>();

    public override void Draw(DrawObject drawObject, float x, float y, float width, float height)
    {
        // Store the object and its position/size for later inspection
        Collected.Add(new DrawObjectInfo
        {
            DrawObject = drawObject,
            X = x,
            Y = y,
            Width = width,
            Height = height
        });
    }
}