using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable macros for the workbook (required for saving as .xlsm)
        workbook.Settings.EnableMacros = true;

        // Add a Forms button control to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Button button = (Button)worksheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);
        button.Text = "Run Macro";

        // Assign the macro name that will be executed when the button is clicked
        button.MacroName = "MyMacro()";

        // Save the workbook as a macro‑enabled file
        workbook.Save("ButtonWithMacro.xlsm", SaveFormat.Xlsm);
    }
}