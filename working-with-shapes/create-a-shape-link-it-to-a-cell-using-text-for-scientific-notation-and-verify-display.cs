// Title: Add a rectangle shape linked to a cell and show its value in scientific notation with the TEXT function using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape text to a worksheet cell using TEXT function | display cell value in scientific notation inside a shape with Aspose.Cells | recalculate formulas so linked shape updates its displayed text in .NET | save workbook with a shape that references a cell formula in Aspose.Cells | set shape text to =TEXT(A1,"0.00E+00") in Aspose.Cells C# example
// Tags: add rectangle shape linked to cell Aspose.Cells | shape text formula using TEXT function Aspose.Cells | scientific notation display in shape Aspose.Cells | formula calculation for shape text Aspose.Cells | persist linked shape when saving workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to cell A1, adds a rectangle shape, assigns the shape's text to the formula =TEXT(A1,"0.00E+00") to render the value in scientific notation, forces formula recalculation so the shape displays the evaluated text, outputs the shape text to the console, and saves the workbook with the linked shape.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Place a numeric value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(12345.6789);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 150, 30);

            // Link the shape's text to cell A1 using the TEXT function for scientific notation
            shape.Text = "=TEXT(A1,\"0.00E+00\")";

            // Recalculate all formulas so the shape shows the evaluated text
            workbook.CalculateFormula();

            // Verify the displayed text by outputting it to the console
            Console.WriteLine("Shape text: " + shape.Text);

            // Save the workbook (optional, demonstrates that the shape is persisted)
            string outputPath = "ShapeLinkedToCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
