using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class AttachMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable macros for the workbook
        workbook.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a CommandButton ActiveX control (acts as a form control)
        Shape buttonShape = worksheet.Shapes.AddActiveXControl(
            ControlType.CommandButton, // type of control
            2,   // upper left row index
            0,   // vertical offset (pixels)
            2,   // upper left column index
            0,   // horizontal offset (pixels)
            100, // width (pixels)
            30   // height (pixels)
        );

        // Set the macro name that will be invoked when the button is clicked
        buttonShape.MacroName = "MyMacro";

        // Optionally set the button caption
        var commandButton = (CommandButtonActiveXControl)buttonShape.ActiveXControl;
        commandButton.Caption = "Run MyMacro";

        // Save the workbook as a macro‑enabled file (XLSM)
        workbook.Save("FormControlWithMacro.xlsm", SaveFormat.Xlsm);
    }
}