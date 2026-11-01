// Title: Create rectangle and oval shapes on a worksheet and save the Excel file using Aspose.Cells for .NET
// Common Searches: how to insert a rectangle shape into an Excel worksheet with Aspose.Cells C# | Aspose.Cells example for adding an oval shape to a worksheet | C# code to add multiple drawing shapes and save workbook using Aspose.Cells | save Excel file with rectangle and oval shapes using Aspose.Cells for .NET | Aspose.Cells drawing API add shapes and export workbook example
// Tags: add rectangle shape Aspose.Cells | insert oval shape worksheet C# | save Excel workbook with drawing shapes Aspose.Cells | Aspose.Cells drawing API C# example | create multiple shapes Excel Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, accesses the first Worksheet, adds a rectangle and an oval shape with specified dimensions and text, then saves the workbook as 'ShapeGroupLockTest.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape
            Shape rectShape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);
            rectShape.Text = "Rectangle";

            // Add an oval shape (used instead of Ellipse)
            Shape ovalShape = sheet.Shapes.AddShape(MsoDrawingType.Oval, 4, 4, 0, 0, 80, 80);
            ovalShape.Text = "Oval";

            // Save the workbook to a file
            string outputPath = "ShapeGroupLockTest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
