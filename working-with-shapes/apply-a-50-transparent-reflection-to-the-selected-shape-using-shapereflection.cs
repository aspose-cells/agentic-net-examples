// Title: Apply a 50% transparent reflection to a rectangle shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook, add a rectangle shape, and assign a 50 % reflection opacity using the Shape.Reflection.Transparency property in Aspose.Cells C#. | Write C# code that generates an .xlsx file containing a shape whose reflection is half transparent via Aspose.Cells. | Use Aspose.Cells for .NET to insert a rectangle into a worksheet and apply a semi‑transparent reflection effect programmatically.
// Common Searches: Aspose.Cells how to make shape reflection 50 percent transparent in C# | C# example for adding rectangle with reflection effect in Excel using Aspose.Cells | Set half transparent reflection on Excel shape via Aspose.Cells API | Programmatically apply reflection opacity to shape in Aspose.Cells .NET
// Tags: Aspose.Cells shape reflection transparency | C# add rectangle shape Aspose.Cells | Reflection.Transparency property example | Excel shape semi transparent effect .NET | Aspose.Cells shape reflection API

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a rectangle shape on the first worksheet, sets the shape's reflection transparency to 50 % using the Reflection.Transparency property, and saves the file as ReflectionShape.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
        var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

        // Apply a 50% transparent reflection to the shape
        // Transparency range: 0 (opaque) to 1 (fully transparent)
        shape.Reflection.Transparency = 0.5;

        // Save the workbook to a file
        workbook.Save("ReflectionShape.xlsx");
    }
}
