// Title: Add an ActiveX ToggleButton to an Excel worksheet using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a ToggleButton ActiveX control with ShapeCollection.AddActiveXControl, set its caption, default unchecked state, enable and visibility flags, and save the file.
// Keywords: Aspose.Cells AddActiveXControl | ToggleButtonActiveXControl C# | ActiveX control Excel .NET | ShapeCollection AddActiveXControl example | set ToggleButton default state
// Common Searches: Aspose.Cells add ToggleButton ActiveX | ShapeCollection AddActiveXControl usage | C# set default value for ToggleButtonActiveXControl | position and size ActiveX control Aspose.Cells | save workbook after inserting ActiveX control
// Developer Intent: Insert a ToggleButton ActiveX control into a worksheet and configure its initial properties.
// Use Cases: Build interactive Excel forms with pre‑placed toggle buttons for user input. | Generate template workbooks that include ready‑to‑use boolean controls. | Automate creation of checklist interfaces by adding multiple ToggleButton controls across rows.
// AI Prompts: Generate C# code that adds a ToggleButton ActiveX control at cell B2, sets the caption to "Enable Feature", and marks it as checked using Aspose.Cells. | Show how to loop from row 2 to row 10 and place an unchecked ToggleButton ActiveX control in column C of each row. | Explain how to read and modify the properties of a saved ToggleButtonActiveXControl when reopening the workbook.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// Demonstrates how to create a workbook, insert a ToggleButton ActiveX control with ShapeCollection.AddActiveXControl, set its caption, default unchecked state, enable and visibility flags, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a ToggleButton ActiveX control.
        // Parameters: control type, top row, top offset (pixels), left column, left offset (pixels), width (pixels), height (pixels)
        Shape shape = worksheet.Shapes.AddActiveXControl(
            ControlType.ToggleButton,
            1,   // upper left row index
            0,   // vertical offset in pixels
            1,   // upper left column index
            0,   // horizontal offset in pixels
            100, // width in pixels
            50   // height in pixels
        );

        // Cast the generic ActiveXControl to a ToggleButtonActiveXControl.
        ToggleButtonActiveXControl toggleButton = (ToggleButtonActiveXControl)shape.ActiveXControl;

        // Set default properties.
        toggleButton.Caption = "Toggle Me";
        toggleButton.Value = CheckValueType.UnChecked; // default state: unchecked
        toggleButton.IsEnabled = true;
        toggleButton.IsVisible = true;

        // Save the workbook.
        workbook.Save("ToggleButtonDemo.xlsx");
    }
}
