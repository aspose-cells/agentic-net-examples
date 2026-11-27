// Title: Remove an ActiveX CommandButton from a worksheet and confirm its removal from the Shapes collection using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove ActiveX CommandButton from worksheet | how to delete an ActiveX control shape with Aspose.Cells .NET | verify that no ActiveX controls remain in a worksheet's Shapes collection using Aspose.Cells | remove ActiveX control from a shape and save the workbook in C# | check for remaining ActiveX shapes after removal with Aspose.Cells
// Tags: remove ActiveX control Aspose.Cells | ActiveX CommandButton shape deletion .NET | verify absence of ActiveX shapes Aspose.Cells | worksheet shapes collection check C# | Aspose.Cells shape ActiveXControl removal

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace RemoveActiveXControlDemo
{
    // The example creates a workbook, adds an ActiveX CommandButton to the first worksheet, saves the file, removes the control from the shape, validates that the shape's ActiveXControl property is null, iterates through all worksheet shapes to ensure no ActiveX controls remain, and finally saves the workbook after removal.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Add an ActiveX CommandButton control to the worksheet.
                // Parameters: ControlType, upper left row, upper left column,
                // upper left row offset, upper left column offset, width, height
                Shape axShape = worksheet.Shapes.AddActiveXControl(
                    ControlType.CommandButton,
                    2, 2,   // upper left row, column
                    0, 0,   // row offset, column offset (pixels)
                    120, 30 // width, height (pixels)
                );

                // Verify that the control was added.
                if (axShape.ActiveXControl == null)
                {
                    Console.WriteLine("Failed to add ActiveX control.");
                    return;
                }
                else
                {
                    Console.WriteLine("ActiveX control added successfully.");
                }

                // Optionally save the workbook before removal for inspection.
                workbook.Save("ActiveXControl_BeforeRemoval.xlsx", SaveFormat.Xlsx);

                // 3. Remove the ActiveX control from the shape.
                axShape.RemoveActiveXControl();

                // 4. Verify that the shape no longer contains an ActiveX control.
                if (axShape.ActiveXControl == null)
                {
                    Console.WriteLine("ActiveX control removed successfully from the shape.");
                }
                else
                {
                    Console.WriteLine("ActiveX control removal failed.");
                }

                // 5. Verify that no shape in the worksheet's Shapes collection contains an ActiveX control.
                bool anyActiveX = false;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.ActiveXControl != null)
                    {
                        anyActiveX = true;
                        // Shape does not expose an Index property; use Type for identification.
                        Console.WriteLine($"Found remaining ActiveX control in shape: Type={shape.Type}");
                    }
                }

                if (!anyActiveX)
                {
                    Console.WriteLine("No ActiveX controls remain in the worksheet's Shapes collection.");
                }
                else
                {
                    Console.WriteLine("There are still ActiveX controls present after removal.");
                }

                // 6. Save the final workbook.
                workbook.Save("ActiveXControl_AfterRemoval.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as 'ActiveXControl_AfterRemoval.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
