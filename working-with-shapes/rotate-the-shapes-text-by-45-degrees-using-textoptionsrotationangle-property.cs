// Title: Rotate a rectangle shape (including its text) by 45 degrees with Aspose.Cells in C#
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet and sets its RotationAngle to 45 degrees using Aspose.Cells. | Show how to rotate a shape and its text by 45 degrees in an Excel file with Aspose.Cells for .NET. | Provide an example of applying a 45-degree rotation to a shape in a workbook and saving it as an .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# rotate shape 45 degrees | How to set shape RotationAngle property in Aspose.Cells | Rotate rectangle shape text in Excel using Aspose.Cells .NET | Example of rotating a shape in Aspose.Cells workbook | C# Aspose.Cells shape rotation angle usage
// Tags: shape rotation angle Aspose.Cells | rotate rectangle shape C# | set shape RotationAngle Excel .NET | Aspose.Cells shape text rotation | apply 45 degree rotation to Excel shape

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a rectangle shape to the first worksheet, assigns the text "Rotated Text" to the shape, rotates the shape (and its text) by 45 degrees using the RotationAngle property, and saves the workbook as RotatedShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                20,  // top offset (pixels)
                20,  // left offset (pixels)
                100, // height (pixels)
                50   // width (pixels)
            );

            // Set the shape's text
            shape.Text = "Rotated Text";

            // Rotate the entire shape (including its text) by 45 degrees
            shape.RotationAngle = 45;

            // Save the workbook to a file
            workbook.Save("RotatedShape.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
