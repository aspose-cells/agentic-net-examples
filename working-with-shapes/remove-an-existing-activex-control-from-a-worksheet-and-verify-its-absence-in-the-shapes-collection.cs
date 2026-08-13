// Title: Remove an ActiveX CommandButton from an Excel worksheet with Aspose.Cells for .NET
// Description: Shows how to add a CommandButton ActiveX control to a worksheet, delete it via Shape.RemoveActiveXControl(), confirm that the Shapes collection no longer contains any ActiveX controls, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | ActiveX control removal | RemoveActiveXControl | worksheet shapes | delete CommandButton | Excel automation | shape collection | verify no ActiveX | programmatic Excel cleanup
// Common Searches: Aspose.Cells remove ActiveX control C# | Delete CommandButton from worksheet using Aspose.Cells | Check for ActiveX controls in Shapes collection Aspose | How to programmatically remove ActiveX from Excel .NET | Aspose.Cells shape.RemoveActiveXControl example
// Developer Intent: The developer needs to delete an existing ActiveX control from a worksheet and ensure it is no longer present in the Shapes collection.
// Use Cases: Remove temporary CommandButton controls added during report generation before publishing the file. | Validate that a shared Excel template is free of ActiveX objects after automated modifications. | Strip all ActiveX controls from a workbook to satisfy security or compatibility policies.
// AI Prompts: Write a C# method that iterates through Worksheet.Shapes and removes any shape containing an ActiveX control using Aspose.Cells. | Provide code to delete a specific ActiveX control by name and return a boolean indicating success. | Generate a snippet that adds an ActiveX control, removes it, and asserts that worksheet.Shapes contains no ActiveX controls.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    // Shows how to add a CommandButton ActiveX control to a worksheet, delete it via Shape.RemoveActiveXControl(), confirm that the Shapes collection no longer contains any ActiveX controls, and save the workbook.
    public class RemoveActiveXControlDemo
    {
        // Entry point required for compilation
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add an ActiveX control (CommandButton) to the worksheet
            Shape shape = worksheet.Shapes.AddActiveXControl(
                ControlType.CommandButton, // type of control
                1,   // upper left row index
                0,   // vertical offset (pixels) from the top row
                1,   // upper left column index
                0,   // horizontal offset (pixels) from the left column
                100, // width (pixels)
                30   // height (pixels)
            );

            // Ensure the control was added
            if (shape.ActiveXControl != null)
                Console.WriteLine("ActiveX control added.");

            // Remove the ActiveX control from the shape
            shape.RemoveActiveXControl();
            Console.WriteLine("ActiveX control removed from the shape.");

            // Verify that no shape in the worksheet contains an ActiveX control
            bool anyActiveX = false;
            foreach (Shape s in worksheet.Shapes)
            {
                if (s.ActiveXControl != null)
                {
                    anyActiveX = true;
                    break;
                }
            }

            Console.WriteLine(anyActiveX
                ? "There is still an ActiveX control present."
                : "No ActiveX controls remain in the Shapes collection.");

            // Save the workbook (optional, demonstrates that the workbook is still valid)
            workbook.Save("RemoveActiveXControlDemo.xlsx");
        }
    }
}
