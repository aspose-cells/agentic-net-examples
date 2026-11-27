// Title: How to read the absolute X and Y pixel positions of a shape named "Logo" in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get pixel coordinates of a specific shape in Excel | retrieve LeftToCorner and TopToCorner values for a named shape with Aspose.Cells | debugging shape location in a workbook using Aspose.Cells .NET API | find absolute position of a shape across all worksheets in Aspose.Cells | C# code to locate shape by name and log its coordinates in an Excel file
// Tags: Aspose.Cells shape position retrieval | C# LeftToCorner TopToCorner properties | Excel shape absolute pixel coordinates | search shape by name across worksheets | debugging shape location Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeCoordinateLogger
{
    // The example loads an Excel workbook, scans every worksheet to locate a shape named "Logo", reads its LeftToCorner and TopToCorner properties (pixel offsets from the worksheet's left and top edges), and writes the absolute X and Y coordinates to the console. It includes error handling for missing files, absent shapes, and unexpected exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains the shape named "Logo"
            const string workbookPath = "input.xlsx";

            // Verify that the file exists before attempting to load it
            if (!System.IO.File.Exists(workbookPath))
            {
                Console.Error.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Assume the shape could be on any worksheet; search all worksheets
                Shape logoShape = null;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        if (string.Equals(shape.Name, "Logo", StringComparison.OrdinalIgnoreCase))
                        {
                            logoShape = shape;
                            break;
                        }
                    }
                    if (logoShape != null)
                        break;
                }

                if (logoShape == null)
                {
                    Console.Error.WriteLine("Debug: Shape named \"Logo\" was not found in any worksheet.");
                    return;
                }

                // Retrieve absolute X (horizontal) and Y (vertical) coordinates.
                // LeftToCorner = horizontal offset from the left border of the worksheet (pixels)
                // TopToCorner  = vertical offset from the top border of the worksheet (pixels)
                int absoluteX = logoShape.LeftToCorner;
                int absoluteY = logoShape.TopToCorner;

                // Log the coordinates for debugging
                Console.WriteLine("Debug: Retrieved coordinates for shape \"Logo\":");
                Console.WriteLine($"       Absolute X (LeftToCorner) = {absoluteX} pixels");
                Console.WriteLine($"       Absolute Y (TopToCorner)  = {absoluteY} pixels");
            }
            catch (Exception ex)
            {
                // Log any unexpected exceptions
                Console.Error.WriteLine($"Exception occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
