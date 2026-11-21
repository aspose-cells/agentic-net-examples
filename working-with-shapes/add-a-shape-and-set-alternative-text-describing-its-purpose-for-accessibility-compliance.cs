// Title: Add a rectangle shape with alternative text for accessibility in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape with alt text in Aspose.Cells C# | setting alternative text for worksheet shapes using Aspose.Cells .NET | Aspose.Cells example for accessibility: add shape and alt description | C# code to insert a shape and define its alternative text in an Excel file | Aspose.Cells add shape and set accessibility properties programmatically
// Tags: add rectangle shape Aspose.Cells | set shape alternative text .NET | Excel accessibility shape Aspose.Cells | Aspose.Cells drawing API rectangle | worksheet shape alt text C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a rectangle shape on the first worksheet, assigns an alternative text description for accessibility, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                2,   // upper left row
                2,   // upper left column
                0,   // top offset (in points)
                0,   // left offset (in points)
                100, // height (in points)
                50   // width (in points)
            );

            // Set alternative text for accessibility compliance
            shape.AlternativeText = "Rectangle indicating the sales target area";

            // Save the workbook to a file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
