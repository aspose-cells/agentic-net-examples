using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeReplaceDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a placeholder text into a cell
        worksheet.Cells["A1"].PutValue("{{PLACEHOLDER}}");

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 50);

        // Set the shape's displayed text to the same placeholder
        shape.Text = "{{PLACEHOLDER}}";

        // Link the shape to the cell containing the placeholder
        // The linked cell is "$A$1". The two boolean flags indicate whether to set the linked cell as a formula and whether to update the shape immediately.
        shape.SetLinkedCell("$A$1", false, false);

        // Replace the placeholder in the workbook with the desired value
        workbook.Replace("{{PLACEHOLDER}}", "Replaced Text");

        // Update the shape so it reflects the new linked cell value
        shape.UpdateSelectedValue();

        // Verify that the shape now shows the replaced text
        Console.WriteLine("Shape text after replace: " + shape.Text);

        // Save the workbook
        workbook.Save("ShapeReplaceDemo.xlsx");
    }
}