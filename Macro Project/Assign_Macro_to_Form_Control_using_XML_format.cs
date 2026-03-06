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

        // Add a Forms button control to the worksheet
        // Parameters: upper left row, upper left column, top offset (pixels), left offset (pixels), width (pixels), height (pixels)
        Button button = (Button)worksheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);
        button.Text = "Run Macro";

        // Assign the macro name to the button using the MacroName property
        button.MacroName = "MyMacro()";

        // (Optional) Define Ribbon XML if you need a custom ribbon UI
        workbook.RibbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"My Tab\">" +
            "        <group id=\"customGroup\" label=\"My Group\">" +
            "          <button id=\"customButton\" label=\"Run Macro\" onAction=\"MyMacro\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Save the workbook as a macro‑enabled file
        workbook.Save("MacroButton.xlsm");
    }
}