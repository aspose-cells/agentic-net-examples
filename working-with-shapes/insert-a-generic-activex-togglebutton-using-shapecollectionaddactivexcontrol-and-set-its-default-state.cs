using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a ToggleButton ActiveX control (AddActiveXControl rule)
        // Parameters: ControlType, topRow, top offset, leftColumn, left offset, width, height
        Shape shape = worksheet.Shapes.AddActiveXControl(
            ControlType.ToggleButton, // type of control
            1,   // upper left row index
            0,   // vertical offset in pixels
            1,   // upper left column index
            0,   // horizontal offset in pixels
            100, // width in pixels
            30   // height in pixels
        );

        // Cast to the specific ToggleButton control
        ToggleButtonActiveXControl toggleButton = (ToggleButtonActiveXControl)shape.ActiveXControl;

        // Set the default state (Value property) to Checked
        toggleButton.Value = CheckValueType.Checked;

        // Optional: set a caption for visual reference
        toggleButton.Caption = "Toggle";

        // Save the workbook (lifecycle rule)
        workbook.Save("ToggleButtonDemo.xlsx");
    }
}