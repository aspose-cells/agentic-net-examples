// Title: How to link a textbox shape to a cell containing a CHAR formula and verify the displayed character using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, write =CHAR(169) into cell B1, add a textbox shape, set its Text property to "=B1", calculate all formulas, and save the file as an .xlsx workbook. | Replace the CHAR formula in the linked cell with UNICHAR(0x00A9), recalculate the workbook, and print the shape's Text value to the console to confirm the correct symbol. | Add several textbox shapes, each linked to different cells that contain CHAR formulas (e.g., CHAR(174), CHAR(176)), run formula calculation, and programmatically verify that each shape displays the expected special character.
// Common Searches: Aspose.Cells .NET link textbox shape to a cell formula that uses CHAR | display copyright symbol in Excel shape using Aspose.Cells | bind shape text to a formula cell and evaluate it with Aspose.Cells C# | how to calculate CHAR function for a linked shape in Aspose.Cells workbook | verify shape text after linking to a cell with special character in C#
// Tags: shape-to-cell formula linking Aspose.Cells | display CHAR result in shape text | bind shape text to formula cell .NET | calculate formulas for linked shapes | save workbook with linked shape xlsx

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a workbook, inserts a CHAR(169) formula into cell B1, adds a textbox shape whose Text property is linked to that cell, calculates formulas so the character is evaluated, prints the shape's text, and saves the workbook as ShapeLinkedToCharCell.xlsx.
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

            // Put a formula in cell B1 that uses CHAR to display a special character (©)
            Cell targetCell = sheet.Cells["B1"];
            targetCell.Formula = "CHAR(169)";

            // Add a textbox shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = sheet.Shapes.AddTextBox(2, 0, 2, 0, 200, 50);

            // Link the shape's text to the cell containing the CHAR formula
            shape.Text = "=B1";

            // Calculate all formulas so the CHAR function is evaluated
            workbook.CalculateFormula();

            // Verify that the shape now displays the evaluated character
            Console.WriteLine("Shape text after linking and calculation: " + shape.Text);

            // Save the workbook
            workbook.Save("ShapeLinkedToCharCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
