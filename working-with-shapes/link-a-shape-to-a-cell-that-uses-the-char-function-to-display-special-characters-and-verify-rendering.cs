// Title: Link a textbox shape to a cell containing a CHAR formula and verify the displayed character using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, sets cell A1 to =CHAR(169), calculates formulas, adds a textbox shape over A1, assigns the shape's Text property from the cell value, and confirms the shape text matches the evaluated character. | Write a method that positions a TextBox shape with PlacementType.MoveAndSize on a specific cell, binds its text to the cell's calculated result, and returns a boolean indicating whether the binding succeeded. | Extend the example to iterate over a collection of cells with different CHAR codes, create corresponding textbox shapes, link each shape's text to its cell, and log any mismatches.
// Common Searches: asp.net link textbox shape to cell formula result aspnet | aspocells set shape text from CHAR function cell | c# verify shape displays special character from Excel formula using Aspose.Cells | how to use PlacementType.MoveAndSize to bind shape to cell in Aspose.Cells | example linking multiple textbox shapes to CHAR formulas in Aspose.Cells
// Tags: Aspose.Cells link textbox shape to cell value | PlacementType.MoveAndSize for shape over cell | use CHAR function with Aspose.Cells formula | validate shape text against calculated cell | C# bind shape text to special character cell

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, writes =CHAR(169) into A1, calculates the formula, adds a textbox shape positioned over the cell with MoveAndSize placement, sets the shape's Text property to the cell's rendered character, verifies the shape text matches the cell value, and saves the workbook as ShapeLinkedToCharFunction.xlsx.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell A1 that uses CHAR to display a special character (©)
            Cell targetCell = sheet.Cells["A1"];
            targetCell.Formula = "=CHAR(169)";

            // Calculate formulas so the cell contains the actual character
            workbook.CalculateFormula();

            // Add a textbox shape to the worksheet (row, column, top offset, left offset, height, width)
            TextBox shape = sheet.Shapes.AddTextBox(0, 0, 0, 0, 30, 200);

            // Position the shape over cell A1
            shape.Placement = PlacementType.MoveAndSize;
            shape.UpperLeftRow = 0;
            shape.UpperLeftColumn = 0;
            // Offsets are already set via AddTextBox parameters; no need for separate properties

            // Link the shape's text to the value of cell A1
            shape.Text = targetCell.StringValue;

            // Verify that the shape's text matches the cell's rendered value
            if (shape.Text == targetCell.StringValue)
            {
                Console.WriteLine("Shape successfully linked. Rendered character: " + shape.Text);
            }
            else
            {
                Console.WriteLine("Link verification failed.");
            }

            // Define output file path
            string outputPath = "ShapeLinkedToCharFunction.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
