using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert an ActiveX CheckBox control with an initial size and position
        Shape shape = worksheet.Shapes.AddActiveXControl(
            ControlType.CheckBox, // type of control
            2, 0,                 // upper left row index and vertical offset (pixels)
            2, 0,                 // upper left column index and horizontal offset (pixels)
            100, 30);             // initial width and height (pixels)

        // Set the desired position (in pixels) after insertion
        shape.Top = 50;    // distance from the top of the worksheet (pixels)
        shape.Left = 150;  // distance from the left of the worksheet (pixels)

        // Set the desired size (in pixels) after insertion
        shape.Width = 200; // width in pixels
        shape.Height = 40; // height in pixels

        // Example: link the control to a cell (optional)
        shape.SetLinkedCell("B5", false, false);

        // Save the workbook
        workbook.Save("ActiveXControlPosition.xlsx");
    }
}