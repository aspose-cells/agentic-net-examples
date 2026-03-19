using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssociateMacroWithFormControl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable macros for the workbook
        workbook.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a button (form control) to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        Button button = sheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

        // Associate a macro with the button
        button.MacroName = "MyMacro()";

        // Optional: set alternative text for accessibility
        button.AlternativeText = "Runs MyMacro";

        // Save the workbook as a macro‑enabled file
        workbook.Save("MacroButtonDemo.xlsm", SaveFormat.Xlsm);
    }
}