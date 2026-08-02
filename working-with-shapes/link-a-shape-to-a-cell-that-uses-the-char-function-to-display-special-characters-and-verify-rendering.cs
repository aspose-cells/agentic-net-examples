// Title: C# – Link a Shape to a CHAR‑Formula Cell and Verify Unicode Output with Aspose.Cells
// Description: Creates a workbook, writes =CHAR(9731) in A1 to generate a snowflake, adds a rectangle shape, links it to A1 using SetLinkedCell, updates the shape with UpdateSelectedValue, reads shape.Text to confirm the Unicode character, and saves the file as ShapeLinkedToCharFunction.xlsx.
// Keywords: Aspose.Cells C# shape linking | SetLinkedCell CHAR function | Unicode character in shape | UpdateSelectedValue Aspose.Cells | retrieve shape text | rectangle shape linked cell | Excel shape formula result
// Common Searches: Aspose.Cells link shape to cell with CHAR formula | display Unicode symbol in linked shape C# | how to use SetLinkedCell and UpdateSelectedValue | verify shape text after linking to cell | C# example linking rectangle to cell formula
// Developer Intent: Demonstrate how to bind a worksheet shape to a cell that contains a CHAR formula and confirm that the shape shows the resulting Unicode character.
// Use Cases: Dynamic icons that reflect formula‑generated symbols | Report templates where shapes act as placeholders for special characters | Automated updates of shape captions when underlying cell formulas change
// AI Prompts: Generate C# code using Aspose.Cells to link a rectangle shape to a cell with a CHAR formula and output the shape's text. | Explain the interaction between SetLinkedCell and UpdateSelectedValue for displaying formula results in a shape. | Provide troubleshooting steps when a linked shape does not show the expected Unicode character.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes =CHAR(9731) in A1 to generate a snowflake, adds a rectangle shape, links it to A1 using SetLinkedCell, updates the shape with UpdateSelectedValue, reads shape.Text to confirm the Unicode character, and saves the file as ShapeLinkedToCharFunction.xlsx.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a cell with the CHAR function to display a special character (e.g., snowflake)
            Cell linkedCell = sheet.Cells["A1"];
            linkedCell.Formula = "=CHAR(9731)"; // Unicode snowflake character

            // Add a rectangle shape (acts as a textbox) to the worksheet
            // Parameters: drawing type, upper left row, upper left column,
            // upper left row offset, upper left column offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // Link the shape to the cell containing the CHAR formula
            shape.SetLinkedCell("A1", true, true);
            // Update the shape's displayed value based on the linked cell
            shape.UpdateSelectedValue();

            // Retrieve the text from the shape (use the Text property)
            string shapeText = shape.Text;
            Console.WriteLine("Shape text (should be the special character): " + shapeText);

            // Save the workbook to a file
            string outputPath = "ShapeLinkedToCharFunction.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
