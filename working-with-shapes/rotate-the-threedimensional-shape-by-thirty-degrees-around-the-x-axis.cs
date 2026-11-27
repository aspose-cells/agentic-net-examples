// Title: Rotate a rectangle shape 30° around the X‑axis using Aspose.Cells ThreeDFormat in C#
// Common Searches: C# Aspose.Cells rotate shape around X axis using ThreeDFormat | How to set RotationX for a 3‑D rectangle in Aspose.Cells .NET | Example of applying 30 degree X‑axis rotation to a shape with Aspose.Cells | Aspose.Cells ThreeDFormat extrusion and X‑axis rotation tutorial | Save a workbook with a rotated 3‑D shape as XLSX using Aspose.Cells
// Tags: Aspose.Cells ThreeDFormat RotationX property | C# rectangle shape 3D extrusion Aspose.Cells | X‑axis rotation of 3D shape Aspose.Cells | apply lighting and bevel to shape Aspose.Cells | save rotated 3D shape to XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace RotateThreeDShapeDemo
{
    // The sample creates a new workbook, adds a rectangle shape, configures its ThreeDFormat (extrusion height, material, lighting, bevel), sets RotationX to 30 degrees to rotate the shape around the X‑axis, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: drawing type, upper left row, upper left column, top, left, height, width
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    5,   // upper left row
                    2,   // upper left column
                    50,  // top (pixels)
                    20,  // left (pixels)
                    150, // height (pixels)
                    200  // width (pixels)
                );

                // Give the shape some visible text
                shape.Text = "3‑D Rotated Shape";

                // Access the ThreeDFormat of the shape
                ThreeDFormat threeD = shape.ThreeDFormat;

                // Set basic 3‑D properties so the effect is visible
                threeD.ExtrusionHeight = 30;          // give the shape depth
                threeD.Material = PresetMaterialType.Plastic;
                threeD.Lighting = LightRigType.ThreePoint;
                threeD.LightingDirection = LightRigDirectionType.Top;
                threeD.TopBevelType = BevelType.SoftRound;
                threeD.TopBevelWidth = 10;
                threeD.TopBevelHeight = 10;

                // Rotate the shape 30 degrees around the X‑axis
                threeD.RotationX = 30.0;

                // Optionally, rotate around Y and Z to see combined effect
                // threeD.RotationY = 0;
                // threeD.RotationZ = 0;

                // Save the workbook to an XLSX file
                string outputPath = "RotatedThreeDShape.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"Shape RotationX = {threeD.RotationX} degrees.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the 3‑D rotated shape:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
