// Title: How to bind a rectangle shape’s text to a cell containing a REPT formula using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, writes =REPT("*",10) into cell A1, adds a rectangle shape, and sets the shape's Text property to =A1 so the shape shows the repeated asterisks. | Write an Aspose.Cells example that links a shape's displayed text to the result of a REPT function in a worksheet cell by assigning a formula to Shape.Text.
// Common Searches: Aspose.Cells C# bind shape text to cell formula REPT | display REPT function output inside a rectangle shape using Aspose.Cells | link Excel shape to cell value with Aspose.Cells .NET example | set shape text to =A1 after using REPT in a worksheet cell
// Tags: Aspose.Cells shape text formula binding | C# rectangle shape REPT function | link shape to cell value Aspose.Cells | Excel shape text from cell formula .NET | Aspose.Cells create shape linked to REPT result

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, assigns the REPT("*",10) formula to cell A1, adds a rectangle shape, links the shape's text to cell A1 by setting Shape.Text to "=A1", and saves the file as LinkedShapeWithRept.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell A1 that repeats a character using REPT function
            // Example: repeat the asterisk (*) 10 times
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=REPT(\"*\",10)";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row (zero-based)
                0,    // upper left column (zero-based)
                0,    // top offset in points
                0,    // left offset in points
                100,  // height in points
                200   // width in points
            );

            // Link the shape's text to cell A1 so it displays the repeated pattern
            // In Aspose.Cells, linking is achieved by setting the shape's text to a formula referencing the cell.
            shape.Text = "=A1";

            // Save the workbook to a file
            string outputPath = "LinkedShapeWithRept.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
