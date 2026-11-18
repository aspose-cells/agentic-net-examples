// Title: Insert a rectangle shape with a clickable hyperlink into an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to add a clickable rectangle shape to an Excel worksheet with Aspose.Cells C# | Aspose.Cells C# example for setting a hyperlink on a shape | create an Excel shape that opens a website when clicked using Aspose.Cells | C# code to attach a URL to a rectangle drawing in an .xlsx file with Aspose | Aspose.Cells add shape and hyperlink programmatically in .NET
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells shape hyperlink .NET | C# insert Excel drawing with URL | Excel shape click-through using Aspose | Aspose.Cells hyperlink on drawing object

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, adds a rectangle shape to the first worksheet, sets its display text, assigns an external URL and screen tip via the shape's Hyperlink property, and saves the file as ShapeWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // Parameters: shape type, upper-left row, upper-left column, top offset, left offset, height, width.
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 200);

            // Set the display text of the shape.
            shape.Text = "Click Here";

            // Assign a hyperlink to the shape using the read‑only Hyperlink object.
            shape.Hyperlink.Address = "https://www.example.com";
            shape.Hyperlink.ScreenTip = "Example Site";

            // Save the workbook to a file.
            string outputPath = "ShapeWithHyperlink.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
