// Title: Add a ToggleButton ActiveX control to an Excel worksheet with Aspose.Cells for .NET
// Description: Creates a new workbook, inserts a generic ToggleButton ActiveX control via ShapeCollection.AddActiveXControl, casts it to ToggleButtonActiveXControl, sets caption, default unchecked state, single‑state mode, enabled and visible flags, and saves the file.
// Keywords: Aspose.Cells | AddActiveXControl | ToggleButton ActiveX | ShapeCollection | C# Excel automation | default ToggleButton state | Enable/Disable ActiveX | Excel workbook template
// Common Searches: Aspose.Cells add ToggleButton ActiveX control | ShapeCollection AddActiveXControl example C# | Set default value for ToggleButtonActiveXControl | Configure visibility of ActiveX control in Aspose.Cells | How to insert generic ActiveX controls into Excel with Aspose
// Developer Intent: Insert a ToggleButton ActiveX control into a worksheet and set its initial properties programmatically.
// Use Cases: Build an interactive Excel form where a pre‑configured ToggleButton records a yes/no choice before data entry. | Generate a template workbook that includes a disabled ToggleButton which becomes enabled after a specific cell meets a condition. | Create a dashboard that uses ToggleButton controls to switch chart series on and off without manual user interaction.
// AI Prompts: Write C# code using Aspose.Cells to place a ToggleButton ActiveX control at cell B2 with a custom caption and set it to the checked state. | Generate a method that adds multiple ToggleButton ActiveX controls to a worksheet and links each button's Value property to a corresponding cell. | Explain how to retrieve an existing ToggleButtonActiveXControl from a saved workbook and modify its IsEnabled and IsVisible properties via Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// Creates a new workbook, inserts a generic ToggleButton ActiveX control via ShapeCollection.AddActiveXControl, casts it to ToggleButtonActiveXControl, sets caption, default unchecked state, single‑state mode, enabled and visible flags, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a ToggleButton ActiveX control at row 1, column 1 with size 100x30 pixels
        Shape shape = sheet.Shapes.AddActiveXControl(
            ControlType.ToggleButton, // control type
            1,   // upper left row index
            0,   // vertical offset in pixels
            1,   // upper left column index
            0,   // horizontal offset in pixels
            100, // width in pixels
            30   // height in pixels
        );

        // Cast the generic ActiveXControl to ToggleButtonActiveXControl
        ToggleButtonActiveXControl toggle = (ToggleButtonActiveXControl)shape.ActiveXControl;

        // Set default properties
        toggle.Caption = "Toggle Me";
        toggle.Value = CheckValueType.UnChecked; // default state unchecked
        toggle.IsTripleState = false;
        toggle.IsEnabled = true;
        toggle.IsVisible = true;

        // Save the workbook
        workbook.Save("ToggleButtonDemo.xlsx");
    }
}
