using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class BindMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a Forms button control to the worksheet
        // Parameters: upper-left row, upper-left column, row offset, column offset, width, height
        Button button = (Button)worksheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

        // Set the macro that will be executed when the button is clicked
        button.MacroName = "MyMacro()";

        // Optional: set the button caption
        button.Text = "Run Macro";

        // Save the workbook in DIF format (the macro binding is stored in the shape)
        workbook.Save("ButtonWithMacro.dif");
    }
}