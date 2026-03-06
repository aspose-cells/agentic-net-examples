using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable macros in the workbook
        workbook.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a Forms button control to the worksheet
        // Parameters: upper left row, upper left column, top offset (pixels), left offset (pixels), width (pixels), height (pixels)
        Button button = (Button)sheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

        // Set the macro name that will be executed when the button is clicked
        button.MacroName = "MyMacro";

        // Optional: set the button caption
        button.Text = "Run Macro";

        // Save the workbook in macro‑enabled XLSM format
        workbook.Save("ButtonWithMacro.xlsm", SaveFormat.Xlsm);
    }
}