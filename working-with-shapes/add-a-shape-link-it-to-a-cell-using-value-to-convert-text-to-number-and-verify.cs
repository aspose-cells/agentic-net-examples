// Title: Create a rectangle shape, bind its text to a numeric cell value, and verify the link using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle shape and set its text from a cell value | Convert cell text to a number and assign it to shape text using Aspose.Cells | Validate that a shape's text matches the numeric value of a worksheet cell in C# | Programmatically link a shape to a worksheet cell value with Aspose.Cells | Example of binding shape text to a cell's numeric conversion in Aspose.Cells for .NET
// Tags: rectangle shape creation Aspose.Cells C# | binding shape text to worksheet cell value | cell string to double conversion Aspose.Cells | shape text verification against cell numeric value | Aspose.Cells shape-to-cell value linking example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new workbook, write a text string to cell A1, add a rectangle shape, convert the cell text to a double, set the shape's text to the numeric value, verify the shape text matches the cell value, and save the file as ShapeLink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a text value in cell A1 (e.g., "123")
            sheet.Cells["A1"].PutValue("123");

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top (in points)
                100,  // left (in points)
                50,   // height (in points)
                100   // width (in points)
            );

            // Convert the text in A1 to a number
            double numericValue = Convert.ToDouble(sheet.Cells["A1"].StringValue);

            // Set the shape's text to the numeric value
            shape.Text = numericValue.ToString();

            // Verify that the shape's text matches the numeric representation of the cell's text
            if (double.TryParse(shape.Text, out double result) && result == numericValue)
            {
                Console.WriteLine("Shape linked correctly. Value: " + result);
            }
            else
            {
                Console.WriteLine("Verification failed.");
            }

            // Save the workbook to a file
            workbook.Save("ShapeLink.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
