// Title: Apply top lighting direction to a shape’s 3‑D format and handle missing specular‑power property using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set LightRigDirectionType.Top for a shape | how to apply top lighting direction in Aspose.Cells ThreeDFormat | C# example for configuring shape lighting direction and extrusion in Aspose.Cells | Aspose.Cells shape specular power not available workaround | set metal preset material and extrusion height for a shape in Aspose.Cells
// Tags: Aspose.Cells ThreeDFormat lighting direction top | C# LightRigDirectionType Top usage | Aspose.Cells shape extrusion height example | Aspose.Cells preset material metal application | Aspose.Cells specular power property limitation

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsThreeDLightingDemo
{
    // The sample creates a workbook, adds a rectangle shape, accesses its ThreeDFormat, sets the LightRigType to ThreePoint, applies a top lighting direction, notes that the specular‑power property is unavailable, adds extrusion height and a metal preset material, and saves the file as an XLSX workbook.
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
                    2,   // upper left row
                    2,   // upper left column
                    100, // top (pixels)
                    100, // left (pixels)
                    150, // height (pixels)
                    250  // width (pixels)
                );

                // Access the 3‑D format of the shape
                ThreeDFormat threeD = shape.ThreeDFormat;

                // Set a light rig type (required for lighting direction to have effect)
                threeD.Lighting = LightRigType.ThreePoint;

                // Apply the lighting direction "Top"
                threeD.LightingDirection = LightRigDirectionType.Top;

                // SpecularPower property does not exist in Aspose.Cells; this line is omitted.

                // Optional: make the effect visible by adding some extrusion and material
                threeD.ExtrusionHeight = 30;
                threeD.Material = PresetMaterialType.Metal;

                // Save the workbook
                string outputPath = "ThreeDLightingTopSpecular80.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
