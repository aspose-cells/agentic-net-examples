// Title: Remove an ActiveX CommandButton from an Excel worksheet using Aspose.Cells for .NET
// Description: C# example that creates a workbook, inserts an ActiveX CommandButton, saves the file, removes the control with Shape.RemoveActiveXControl(), checks the worksheet's Shapes collection to confirm the control is gone, and saves the updated workbook.
// Keywords: Aspose.Cells | C# | remove ActiveX control | delete CommandButton | Shape.RemoveActiveXControl | Excel shapes collection | verify ActiveX removal | Aspose.Cells .NET example
// Common Searches: how to delete an ActiveX CommandButton with Aspose.Cells | remove ActiveX control from Excel worksheet C# | check for remaining ActiveX controls in Aspose.Cells | Aspose.Cells shape.RemoveActiveXControl usage | C# code to remove ActiveX controls from workbook
// Developer Intent: Delete an existing ActiveX control from a worksheet and confirm that it no longer appears in the Shapes collection.
// Use Cases: Clean up temporary ActiveX buttons before publishing a workbook. | Programmatically strip interactive controls after automated processing. | Validate that a template workbook is free of ActiveX objects prior to distribution.
// AI Prompts: Write C# code using Aspose.Cells to remove all ActiveX controls from a worksheet and verify the removal by iterating the Shapes collection. | Show an example that adds an ActiveX CommandButton, removes it with Shape.RemoveActiveXControl(), and prints a message indicating whether any ActiveX controls remain.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, inserts an ActiveX CommandButton, saves the file, removes the control with Shape.RemoveActiveXControl(), checks the worksheet's Shapes collection to confirm the control is gone, and saves the updated workbook.
    public class RemoveActiveXControlDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add an ActiveX CommandButton control to the worksheet
                Shape shape = worksheet.Shapes.AddActiveXControl(
                    ControlType.CommandButton, // type of control
                    1,   // upper left row index
                    0,   // vertical offset (pixels) from the top of the row
                    1,   // upper left column index
                    0,   // horizontal offset (pixels) from the left of the column
                    100, // width (pixels)
                    30   // height (pixels)
                );

                // Ensure the control was added
                if (shape.ActiveXControl == null)
                {
                    Console.WriteLine("Failed to add ActiveX control.");
                    return;
                }

                // Save the workbook before removal (optional)
                string beforePath = Path.Combine(Directory.GetCurrentDirectory(), "BeforeRemoval.xlsx");
                workbook.Save(beforePath, SaveFormat.Xlsx);

                // Remove the ActiveX control from the shape
                shape.RemoveActiveXControl();

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
                    ? "ActiveX control still present in Shapes collection."
                    : "ActiveX control successfully removed; no ActiveX controls in Shapes collection.");

                // Save the workbook after removal
                string afterPath = Path.Combine(Directory.GetCurrentDirectory(), "AfterRemoval.xlsx");
                workbook.Save(afterPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveActiveXControlDemo.Run();
        }
    }
}
