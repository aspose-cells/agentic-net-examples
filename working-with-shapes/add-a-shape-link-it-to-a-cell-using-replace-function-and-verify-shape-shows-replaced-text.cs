// Title: Create a rectangle shape, assign a REPLACE formula linked to cell A1, evaluate it, and save the workbook with Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to the first worksheet, set its Text property to =REPLACE(A1,1,5,"Hi"), call workbook.CalculateFormula(), and output the evaluated shape text. | Replace the shape's Text with a SUBSTITUTE formula that changes "World" to "Universe", recalculate the workbook, and display the updated shape text. | After formula evaluation, save the workbook as ShapeReplaceResult.xlsx and print the full file path.
// Common Searches: how to bind a shape's text to an Excel formula using Aspose.Cells C# | evaluate REPLACE function inside a shape with Aspose.Cells .NET | Aspose.Cells calculate formulas in shapes after adding them | set rectangle shape text to reference cell A1 in C# Aspose.Cells | save workbook after shape formula evaluation Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | shape text REPLACE formula evaluation | calculate workbook formulas Aspose.Cells | save workbook with evaluated shape text .xlsx | link shape text to worksheet cell using formula

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes "HelloWorld" to cell A1, adds a rectangle shape, assigns a REPLACE formula to the shape's Text property, forces formula calculation, reads and prints the evaluated shape text, and saves the result as ShapeReplaceResult.xlsx.
class ShapeReplaceExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put sample text into cell A1
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue("HelloWorld");

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, lower right row, lower right column offset
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 4, 200);

            // Set the shape's text to a REPLACE formula that operates on A1
            shape.Text = "=REPLACE(A1,1,5,\"Hi\")";

            // Calculate all formulas in the workbook so the shape's text is evaluated
            workbook.CalculateFormula();

            // Retrieve the evaluated text from the shape
            string shapeText = shape.Text;
            Console.WriteLine("Shape text after REPLACE: " + shapeText);

            // Save the workbook
            string outputPath = "ShapeReplaceResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
