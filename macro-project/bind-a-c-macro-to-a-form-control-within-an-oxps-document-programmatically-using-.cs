using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class Program
{
    static void Main()
    {
        // Create a new workbook (macro‑enabled format will be used when saving)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a CommandButton ActiveX control to the worksheet
        // Parameters: control type, upper‑left row, upper‑left column,
        // upper‑left pixel offset, lower‑right pixel offset, width, height
        Shape buttonShape = worksheet.Shapes.AddActiveXControl(
            ControlType.CommandButton, 2, 2, 0, 0, 100, 30);

        // Cast to the specific ActiveX control to set its visual properties
        CommandButtonActiveXControl commandButton = (CommandButtonActiveXControl)buttonShape.ActiveXControl;
        commandButton.Caption = "Run Macro";

        // Bind a VBA macro to the button. The macro must exist in the workbook.
        // The MacroName property expects the macro call, e.g., "MyMacro()".
        buttonShape.MacroName = "MyMacro()";

        // Save the workbook as a macro‑enabled file so the macro binding is retained
        workbook.Save("MacroButtonDemo.xlsm");
    }
}