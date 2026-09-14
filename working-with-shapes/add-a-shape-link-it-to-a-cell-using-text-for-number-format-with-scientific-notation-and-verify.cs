// Title: Add a rectangle shape linked to a cell using the TEXT function for scientific notation in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a rectangle shape on the first worksheet, assigns its Text property to =TEXT(A1,"0.00E+00"), forces formula calculation, and saves the file as ShapeLinkedScientific.xlsx. | Write a verification routine that compares the shape's displayed text after calculation with the .NET formatted scientific notation of the cell value and outputs pass/fail messages.
// Common Searches: Aspose.Cells link shape text to a cell with TEXT function scientific notation | C# add rectangle shape that shows cell value in scientific format using Aspose.Cells | verify that shape text matches the evaluated formula in Aspose.Cells .NET | save workbook containing a shape bound to a cell formula Aspose.Cells example
// Tags: rectangle shape creation Aspose.Cells | shape text formula binding Aspose.Cells | scientific notation formatting in shape | formula evaluation for shape text | shape verification against .NET format

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape, bind its text to cell A1 with the TEXT function using scientific notation, calculate formulas so the shape displays the evaluated result, verify the displayed value matches .NET formatting, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a numeric value into cell A1
        worksheet.Cells["A1"].PutValue(123456789);

        // Add a rectangle shape to the worksheet
        // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
        Shape shape = worksheet.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            2,   // upper left row
            2,   // upper left column
            0,   // top offset (pixels)
            0,   // left offset (pixels)
            100, // height (pixels)
            200  // width (pixels)
        );

        // Link the shape's text to cell A1 using the TEXT function with scientific notation
        shape.Text = "=TEXT(A1,\"0.00E+00\")";

        // Calculate all formulas so the shape displays the evaluated result
        workbook.CalculateFormula();

        // Verify that the shape displays the expected scientific notation
        string expected = 123456789.ToString("0.00E+00"); // .NET formatting for comparison
        string actual = shape.Text; // After calculation this holds the displayed text

        Console.WriteLine($"Expected: {expected}");
        Console.WriteLine($"Actual  : {actual}");

        if (actual == expected)
        {
            Console.WriteLine("Verification passed: Shape text matches scientific notation.");
        }
        else
        {
            Console.WriteLine("Verification failed: Shape text does not match expected format.");
        }

        // Save the workbook
        workbook.Save("ShapeLinkedScientific.xlsx");
    }
}
