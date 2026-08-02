using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RenameButtonDemo
{
    static void Main()
    {
        // Load an existing workbook that contains a form control button
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // The current name of the button we want to rename
        string oldButtonName = "MyButton";

        // Locate the button shape by its name
        Button button = null;
        foreach (Shape shape in worksheet.Shapes)
        {
            if (shape is Button && shape.Name == oldButtonName)
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
            button.Name = "NewButtonName";

            // Reassign the macro name to ensure it remains unchanged
            button.MacroName = existingMacro;
        }
        else
        {
            Console.WriteLine($"Button with name '{oldButtonName}' not found.");
        }

        // Save the workbook with the updated button name
        workbook.Save("output.xlsx");
    }
}