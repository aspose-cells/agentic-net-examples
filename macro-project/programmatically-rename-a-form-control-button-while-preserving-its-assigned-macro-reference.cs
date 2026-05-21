using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RenameButtonDemo
{
    static void Main()
    {
        // Load the workbook that contains the button
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet where the button resides
        Worksheet worksheet = workbook.Worksheets[0];

        // Locate the button by its current name
        Button button = null;
        foreach (Shape shape in worksheet.Shapes)
        {
            if (shape is Button && shape.Name == "OldButton")
            {
                button = (Button)shape;
                break;
            }
        }

        if (button != null)
        {
            // Store the existing macro reference
            string existingMacro = button.MacroName;

            // Rename the button
            button.Name = "NewButton";

            // Reassign the macro name to ensure it remains linked
            button.MacroName = existingMacro;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}