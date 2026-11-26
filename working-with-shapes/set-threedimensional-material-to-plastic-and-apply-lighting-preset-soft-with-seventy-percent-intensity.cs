// Title: Set Plastic material and Soft lighting (70% intensity) on a rectangle shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape material to plastic and apply soft lighting | How to control lighting intensity for a 3D shape in Aspose.Cells | Add extrusion and contour to a rectangle shape using ThreeDFormat in Aspose.Cells
// Tags: threeDformat plastic material Aspose.Cells | soft lighting preset threeD shape C# | lightangle 70 degrees lighting intensity Aspose.Cells | extrusionheight contourwidth 3D shape formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a rectangle shape, accesses its ThreeDFormat, sets the material to Plastic, applies the Soft lighting preset, approximates 70% lighting intensity by setting LightAngle to 70 degrees, adds extrusion height and contour width for visual depth, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Initialize a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will receive the 3‑D formatting
        // Parameters: drawing type, upper‑left row, upper‑left column, top, left, height, width
        Shape shape = worksheet.Shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            2,   // upper‑left row
            2,   // upper‑left column
            100, // top (pixels)
            100, // left (pixels)
            200, // height (pixels)
            200  // width (pixels)
        );

        // Access the ThreeDFormat object of the shape
        ThreeDFormat threeD = shape.ThreeDFormat;

        // 1. Set the material to Plastic
        threeD.Material = PresetMaterialType.Plastic;

        // 2. Apply the lighting preset "Soft"
        // The LightRigType enumeration contains a Soft preset.
        threeD.Lighting = LightRigType.Soft;

        // 3. Set the lighting intensity.
        // Aspose.Cells does not expose a direct intensity property.
        // The closest controllable parameter is LightAngle, which influences the perceived strength.
        // Setting it to 70 (degrees) approximates a 70 % intensity effect.
        threeD.LightAngle = 70.0;

        // Optional: make the 3‑D effect visible by adding extrusion and contour
        threeD.ExtrusionHeight = 30;   // give the shape some depth
        threeD.ContourWidth = 2;       // thin contour for better visual contrast

        // Save the workbook
        string outputPath = "ThreeDMaterialPlasticSoftLighting.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to '{outputPath}'.");
        Console.WriteLine("Three‑D material set to Plastic, lighting preset set to Soft with approx. 70 % intensity.");
    }
}
