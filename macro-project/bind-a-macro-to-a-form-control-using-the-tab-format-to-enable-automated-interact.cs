using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class BindMacroToControl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a CommandButton ActiveX control to the worksheet
        // Parameters: control type, upper-left row, column, row offset, column offset, width, height
        Shape shape = sheet.Shapes.AddActiveXControl(
            ControlType.CommandButton,
            2, 2,
            2, 2,
            100, 30);

        // Access the specific ActiveX control to set its caption
        CommandButtonActiveXControl button = (CommandButtonActiveXControl)shape.ActiveXControl;
        button.Caption = "Run Macro";

        // Bind a macro to the control using the TAB format (MacroName property)
        shape.MacroName = "MyMacro()";

        // Save the workbook with the bound macro
        workbook.Save("MacroBoundControl.xlsx");
    }
}