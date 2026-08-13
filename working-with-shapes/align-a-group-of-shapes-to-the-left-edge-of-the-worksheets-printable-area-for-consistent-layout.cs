// Title: Align worksheet shapes to the printable area's left margin with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds several shapes, reads the worksheet's left margin (cm), converts it to points, and sets each shape's LeftToCorner so all shapes line up with the printable area's left edge before saving.
// Keywords: Aspose.Cells shape alignment | C# left margin printable area | LeftToCorner property | convert cm to points Aspose.Cells | position Excel shapes programmatically | page setup margins Aspose.Cells
// Common Searches: how to align all shapes to left printable area Aspose.Cells | set shape left position based on page margin .NET | convert worksheet margin centimeters to points for shape placement | Aspose.Cells align multiple shapes left edge | C# code to position Excel shapes using printable area
// Developer Intent: Position every shape on a worksheet so its left edge matches the left margin of the printable area.
// Use Cases: Design a report template where logos and graphics start at the printable area's left edge for consistent printing. | Standardize watermarks or header images across many sheets after adjusting page margins. | Automate bulk updates of shape positions when the left margin changes, eliminating manual repositioning.
// AI Prompts: Generate C# code using Aspose.Cells to align all shapes to the right margin of the printable area. | Show how to read and modify the top margin to vertically align shapes in an Excel worksheet with Aspose.Cells. | Provide an example that aligns shapes to the printable area while preserving their original size and aspect ratio.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AlignShapesToPrintableArea
{
    // C# example that creates a workbook, adds several shapes, reads the worksheet's left margin (cm), converts it to points, and sets each shape's LeftToCorner so all shapes line up with the printable area's left edge before saving.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a few sample shapes to demonstrate the alignment
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape shape1 = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);
                Shape shape2 = sheet.Shapes.AddOval(5, 3, 0, 0, 120, 60);
                Shape shape3 = sheet.Shapes.AddTextBox(8, 1, 0, 0, 150, 70);

                // Optional: give shapes some visual distinction (commented out due to API differences)
                // shape1.Fill.ForeColor = System.Drawing.Color.LightBlue;
                // shape2.Fill.ForeColor = System.Drawing.Color.LightGreen;
                // shape3.Fill.ForeColor = System.Drawing.Color.LightCoral;

                // Align all shapes to the left edge of the printable area.
                // The printable area starts after the left margin defined in PageSetup.
                // Convert the left margin (centimeters) to points (1 cm = 28.3464567 points)
                // Shape.LeftToCorner expects an integer offset (in points).
                double leftMarginCm = sheet.PageSetup.LeftMargin; // default is 2.54 cm (1 inch)
                int leftMarginPoints = (int)Math.Round(leftMarginCm * 28.3464567);

                // Iterate through all shapes on the worksheet and set their LeftToCorner
                foreach (Shape shp in sheet.Shapes)
                {
                    shp.LeftToCorner = leftMarginPoints;
                }

                // Define output file path
                string outputPath = "AlignedShapes.xlsx";

                // Save the workbook (lifecycle: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
