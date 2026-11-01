// Title: Bind a rectangle shape’s text to a CONCAT formula and verify the result using Aspose.Cells for .NET (C#)
// AI Prompts: Create a rectangle shape on a worksheet and assign its Text property to the formula =CONCAT(A1,B1) with Aspose.Cells in C#. | Run Workbook.CalculateFormula to evaluate the shape’s formula, then read Shape.Text to verify the concatenated output. | Save the workbook to an .xlsx file and print the absolute file path after linking the shape to the cell values.
// Common Searches: asp.net how to bind a shape’s text to a CONCAT formula using Aspose.Cells | c# display concatenated cell values inside a rectangle shape with Aspose.Cells | retrieve evaluated shape text after calling CalculateFormula in Aspose.Cells | link shape to worksheet cells via formula and save as xlsx Aspose.Cells C#
// Tags: Aspose.Cells set shape text to formula | C# rectangle shape CONCAT cell values | calculate workbook formulas Aspose.Cells | read evaluated shape text Aspose.Cells | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, writes "Hello " to A1 and "World!" to B1, adds a rectangle shape, sets its Text to the CONCAT formula referencing those cells, recalculates formulas, reads the shape's displayed text (expected "Hello World!"), and saves the file as ShapeConcat.xlsx.
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

            // Put sample values into cells A1 and B1
            sheet.Cells["A1"].PutValue("Hello ");
            sheet.Cells["B1"].PutValue("World!");

            // Add a rectangle shape to the worksheet
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 100, 30);

            // Set the shape's text to a formula that concatenates A1 and B1
            shape.Text = "=CONCAT(A1,B1)";

            // Recalculate all formulas so the shape displays the result
            workbook.CalculateFormula();

            // Retrieve the displayed text from the shape after calculation
            string displayedText = shape.Text;
            Console.WriteLine("Shape displays: " + displayedText); // Expected: "Hello World!"

            // Save the workbook
            string outputPath = "ShapeConcat.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
