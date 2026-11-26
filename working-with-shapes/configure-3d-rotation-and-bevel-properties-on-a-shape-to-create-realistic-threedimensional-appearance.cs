// Title: Configure 3D rotation, bevel, extrusion, and lighting for a rectangle shape with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells set rectangle shape 3D rotation X Y Z | How to apply top and bottom bevel to a shape using Aspose.Cells .NET | Aspose.Cells add extrusion height and material to a shape in C# | Setting lighting and contour for 3D formatted shape with Aspose.Cells | Create realistic 3D rectangle shape in Excel using Aspose.Cells API
// Tags: Aspose.Cells shape ThreeDFormat rotation | Aspose.Cells rectangle shape bevel configuration | Aspose.Cells extrusion height and color | Aspose.Cells preset material and light rig | Aspose.Cells shape contour styling

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a workbook, inserts a rectangle shape, and uses the shape's ThreeDFormat to define top and bottom bevels, extrusion height and color, X/Y/Z rotations, metal material, three‑point lighting with direction and angle, and a dark‑gray contour, then saves the file as ThreeDShapeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top, left, height, width
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 10, 10, 150, 200);
            shape.Name = "ThreeDShape";
            shape.Text = "3D Demo";

            // Access the shape's 3‑D formatting object
            ThreeDFormat threeD = shape.ThreeDFormat;

            // ----- Bevel settings -----
            // Top bevel
            threeD.TopBevelType = BevelType.SoftRound;
            threeD.TopBevelWidth = 12;   // points
            threeD.TopBevelHeight = 12;  // points

            // Bottom bevel
            threeD.BottomBevelType = BevelType.Divot;
            threeD.BottomBevelWidth = 8;
            threeD.BottomBevelHeight = 8;

            // ----- Extrusion settings -----
            threeD.ExtrusionHeight = 30; // points
            threeD.ExtrusionColor.Color = Color.LightBlue;

            // ----- Rotation settings (degrees) -----
            threeD.RotationX = 30; // rotate around X‑axis
            threeD.RotationY = 20; // rotate around Y‑axis
            threeD.RotationZ = 15; // rotate around Z‑axis

            // ----- Material and lighting for realism -----
            threeD.Material = PresetMaterialType.Metal;
            threeD.Lighting = LightRigType.ThreePoint;
            threeD.LightingDirection = LightRigDirectionType.Top;
            threeD.LightAngle = 45; // angle of the light source

            // ----- Contour (edge) settings -----
            threeD.ContourWidth = 2; // points
            threeD.ContourColor.Argb = Color.DarkGray.ToArgb();

            // Save the workbook
            string outputPath = "ThreeDShapeDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
