using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a Forms button to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        Button button = worksheet.Shapes.AddButton(1, 1, 1, 1, 100, 30);

        // Set basic properties of the button
        button.Name = "MyButton";
        button.Text = "Run Macro";

        // Assign the macro name that will be executed when the button is clicked
        button.MacroName = "MyMacro";

        // Save the workbook in XPS format
        workbook.Save("ButtonWithMacro.xps", SaveFormat.Xps);
    }
}