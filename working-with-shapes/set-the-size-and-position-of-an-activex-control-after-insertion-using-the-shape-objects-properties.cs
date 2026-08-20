// Title: Set Size and Position of an ActiveX Control via Shape Properties in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to insert an ActiveX CheckBox into a worksheet, then use the Shape object's Top, Left, Width, and Height properties (pixel units) to reposition and resize the control, optionally link it to a cell, and save the workbook.
// Keywords: Aspose.Cells ActiveX control position | Aspose.Cells Shape Top property | Aspose.Cells Shape Left property | Aspose.Cells Shape Width Height | C# Aspose.Cells set ActiveX size | AddActiveXControl Aspose.Cells | SetLinkedCell ActiveX Aspose | Excel ActiveX checkbox positioning | pixel based layout Aspose.Cells | .NET Excel shape manipulation
// Common Searches: how to move an ActiveX checkbox in Aspose.Cells C# | change width and height of an inserted ActiveX control using Aspose.Cells | set pixel offset for ActiveX controls with Shape object | link ActiveX control to a cell after repositioning Aspose.Cells | Aspose.Cells AddActiveXControl example
// Developer Intent: Modify the location and dimensions of a newly added ActiveX control by assigning values to the Shape's Top, Left, Width, and Height properties.
// Use Cases: Align a checkbox with a data label by adjusting Shape.Top and Shape.Left after insertion. | Enlarge an ActiveX dropdown to fit longer list items by changing Shape.Width and Shape.Height. | Place a custom button at an exact pixel offset for a tailored UI before linking it to a worksheet cell.
// AI Prompts: Generate C# code that adds an ActiveX button at row 5, column 3 and sets its size to 120 × 35 pixels using Aspose.Cells. | Show how to move an existing ActiveX control to cell D10 with a 10‑pixel left offset by updating Shape.Top and Shape.Left. | Explain the steps to link an ActiveX control to a cell after changing its position with Shape properties in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// Demonstrates how to insert an ActiveX CheckBox into a worksheet, then use the Shape object's Top, Left, Width, and Height properties (pixel units) to reposition and resize the control, optionally link it to a cell, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert an ActiveX CheckBox control with initial size and position
        Shape shape = worksheet.Shapes.AddActiveXControl(
            ControlType.CheckBox, // type of control
            2,    // upper left row index
            0,    // vertical offset (pixels) from the top row
            2,    // upper left column index
            0,    // horizontal offset (pixels) from the left column
            80,   // initial width (pixels)
            30    // initial height (pixels)
        );

        // Set the desired position and size using the Shape object's properties (pixels)
        shape.Top = 50;      // vertical offset from the top row
        shape.Left = 100;    // horizontal offset from the left column
        shape.Width = 150;   // width of the control
        shape.Height = 40;   // height of the control

        // Example: link the control to a cell (optional)
        shape.SetLinkedCell("B2", false, false);

        // Save the workbook
        workbook.Save("ActiveXControlPositionDemo.xlsx");
    }
}
