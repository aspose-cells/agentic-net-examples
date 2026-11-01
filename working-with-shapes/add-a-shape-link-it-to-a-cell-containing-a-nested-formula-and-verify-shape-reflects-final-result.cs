// Title: Create a rectangle shape, bind its text to a cell with a SUM(A1:A3) formula, and verify the displayed value using Aspose.Cells for .NET
// AI Prompts: Add a rectangle shape to a worksheet, set its Text property to "=B1" where B1 contains a SUM(A1:A3) formula, recalculate the workbook, and print both the cell and shape values. | Link a shape's text to a cell that holds a nested formula, invoke Workbook.CalculateFormula, then compare shape.Text with the cell's evaluated result in C# with Aspose.Cells.
// Common Searches: Aspose.Cells how to bind shape text to a cell formula in C# | C# Aspose.Cells display result of SUM formula inside a shape | verify that a linked shape shows the evaluated value of a cell in Aspose.Cells | link rectangle shape to cell B1 containing SUM(A1:A3) using Aspose.Cells for .NET
// Tags: rectangle shape linked to cell formula Aspose.Cells | shape text reference to worksheet cell C# | recalculate formulas before accessing shape text Aspose.Cells | display SUM(A1:A3) result inside shape .NET | validate shape text matches evaluated cell value

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, fills A1‑A3 with numbers, sets a SUM(A1:A3) formula in B1, adds a rectangle shape, links the shape's text to B1 using "=B1", calculates all formulas, retrieves and prints the evaluated cell value and the shape's displayed text, checks that they match, and saves the file as ShapeLinkedToFormula.xlsx.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells A1:A3 with numeric values
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a nested formula in B1: SUM of A1:A3
            sheet.Cells["B1"].Formula = "SUM(A1:A3)";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 5, 0, 0, 100, 200);

            // Link the shape's text to cell B1 using a formula reference
            shape.Text = "=B1";

            // Calculate all formulas so that the shape can reflect the evaluated result
            workbook.CalculateFormula();

            // Retrieve the evaluated value from the cell
            string cellResult = sheet.Cells["B1"].StringValue;

            // Retrieve the text displayed by the shape (should be the same as cellResult)
            string shapeResult = shape.Text;

            // Verify that the shape reflects the final result of the nested formula
            Console.WriteLine($"Cell B1 evaluated value: {cellResult}");
            Console.WriteLine($"Shape linked text: {shapeResult}");
            Console.WriteLine($"Verification: {(cellResult == shapeResult ? "Success" : "Failure")}");

            // Save the workbook (lifecycle rule)
            workbook.Save("ShapeLinkedToFormula.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
