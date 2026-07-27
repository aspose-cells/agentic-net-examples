// Title: Resize and reposition an ActiveX control using Shape properties in AspNet Aspose.Cells
// Description: Demonstrates how to insert an ActiveX CheckBox into a worksheet and then change its location and dimensions by setting the Shape object's Top, Left, Width, and Height properties before saving the workbook.
// Keywords: Aspose.Cells ActiveX resize | Shape.Top Shape.Left Aspose.Cells | C# set ActiveX control size | adjust Excel ActiveX position programmatically | Aspose.Cells .NET shape dimensions
// Common Searches: change size of ActiveX control after adding with Aspose.Cells | move ActiveX checkbox to another cell using Shape properties | set width and height of ActiveX control in C# Aspose.Cells | reposition Excel ActiveX element programmatically
// Developer Intent: Modify the pixel‑based Top, Left, Width, and Height of a Shape that hosts an ActiveX control after it has been added to a worksheet.
// Use Cases: Align a newly added ActiveX button with a custom form layout. | Update the position of a checkbox when rows are inserted or deleted. | Scale a dropdown list to fit the width of merged cells at runtime.
// AI Prompts: Write C# code that adds an ActiveX ListBox at cell C5 with Aspose.Cells and sets its Top, Left, Width, and Height to specific pixel values. | Show how to relocate an existing ActiveX control to column D and increase its height using the Shape object's properties in Aspose.Cells for .NET. | Create an example that inserts three ActiveX controls and aligns them in a row by assigning identical Shape.Top values and incremental Shape.Left offsets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsActiveXExample
{
    // Demonstrates how to insert an ActiveX CheckBox into a worksheet and then change its location and dimensions by setting the Shape object's Top, Left, Width, and Height properties before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert an ActiveX CheckBox control at row 2, column 2 with initial size
            Shape shape = worksheet.Shapes.AddActiveXControl(
                ControlType.CheckBox, // type of control
                1,    // upper left row index (zero‑based)
                0,    // vertical offset in pixels
                1,    // upper left column index (zero‑based)
                0,    // horizontal offset in pixels
                100,  // initial width in pixels
                50    // initial height in pixels
            );

            // Adjust the position and size using the Shape object's properties
            shape.Top = 20;      // vertical offset from the top row (pixels)
            shape.Left = 30;    // horizontal offset from the left column (pixels)
            shape.Width = 150;  // new width (pixels)
            shape.Height = 40;  // new height (pixels)

            // Optionally, access the underlying ActiveX control to set additional properties
            CheckBoxActiveXControl checkBox = (CheckBoxActiveXControl)shape.ActiveXControl;
            checkBox.Caption = "Accept Terms";
            checkBox.IsEnabled = true;

            // Save the workbook
            workbook.Save("ActiveXControlPositionSize.xlsx");
        }
    }
}
