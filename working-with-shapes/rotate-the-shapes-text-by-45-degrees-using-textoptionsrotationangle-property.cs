// Title: Rotate Shape Text 45° with Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape to a worksheet, assign text, and rotate that text 45 degrees using the TextOptions.RotationAngle (or Shape.RotationAngle) property, then save the workbook as an .xlsx file.
// Keywords: Aspose.Cells rotate text | C# shape text rotation | TextOptions RotationAngle example | Excel shape angled text .NET | rotate rectangle label Aspose
// Common Searches: Aspose.Cells rotate shape text 45 degrees | C# set text rotation angle in Excel shape | How to tilt text inside a shape using Aspose.Cells | TextOptions.RotationAngle Aspose.Cells .NET
// Developer Intent: Apply a 45‑degree rotation to the text of a worksheet shape using Aspose.Cells for .NET.
// Use Cases: Create diagonal labels for flow‑chart elements. | Design angled callouts in dashboards or reports. | Fit longer captions into limited space by slanting the text.
// AI Prompts: Write C# code that adds a shape, sets its TextOptions.RotationAngle to 45°, and saves the workbook with Aspose.Cells. | Explain when to use Shape.RotationAngle versus TextOptions.RotationAngle for rotating text in Aspose.Cells. | Show how to rotate the text of multiple shapes by different angles in a single workbook using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a rectangle shape to a worksheet, assign text, and rotate that text 45 degrees using the TextOptions.RotationAngle (or Shape.RotationAngle) property, then save the workbook as an .xlsx file.
    public class RotateShapeTextDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, height, width
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 200, 300, 400);

                // Set the text of the shape
                shape.Text = "Rotated Text";

                // Rotate the shape (including its text) by 45 degrees
                shape.RotationAngle = 45;

                // Define output file path
                string outputPath = "ShapeTextRotated45.xlsx";

                // Save the workbook
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
