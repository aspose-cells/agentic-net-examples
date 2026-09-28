// Title: Rotate a 3D rectangle shape 30° around the X‑axis using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet and sets its ThreeDFormat.RotationX to 30 degrees with Aspose.Cells. | Create an Excel workbook containing a 3‑D formatted shape rotated 30° on the X‑axis via the Aspose.Cells Shape.ThreeDFormat API. | Write a C# program that demonstrates applying X‑axis rotation to a shape's ThreeDFormat property and saving the file using Aspose.Cells.
// Common Searches: Aspose.Cells C# rotate shape 30 degrees on X axis | set ThreeDFormat.RotationX for rectangle shape in Aspose.Cells | example of 3D shape rotation in Excel using Aspose.Cells .NET | how to apply X‑axis rotation to a worksheet shape with Aspose.Cells | C# code to create 3D rectangle shape and rotate it in an Excel file
// Tags: Aspose.Cells shape ThreeDFormat rotation | C# rotate worksheet shape X‑axis | Excel 3D shape rotation Aspose.Cells | set RotationX property Aspose.Cells | create 3D rectangle shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a three‑dimensional rectangle shape to the first worksheet, sets its ThreeDFormat.RotationX to 30 degrees, saves the workbook as RotatedShape.xlsx, and handles any exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a three‑dimensional rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,    // upper‑left row
                    1,    // upper‑left column
                    0,    // upper‑left row offset (pixels)
                    0,    // upper‑left column offset (pixels)
                    150,  // width (pixels)
                    100   // height (pixels)
                );

                // Enable 3‑D formatting by setting rotation (3‑D is applied automatically)
                shape.ThreeDFormat.RotationX = 30;

                // Save the workbook to a file
                string outputPath = "RotatedShape.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
