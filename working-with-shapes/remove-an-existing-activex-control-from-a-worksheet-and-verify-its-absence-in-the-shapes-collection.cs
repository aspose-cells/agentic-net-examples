using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
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
                    ? "There is still an ActiveX control present in the worksheet."
                    : "No ActiveX controls remain in the worksheet.");

                // Save the workbook (optional, just to demonstrate lifecycle)
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RemoveActiveXControlDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}