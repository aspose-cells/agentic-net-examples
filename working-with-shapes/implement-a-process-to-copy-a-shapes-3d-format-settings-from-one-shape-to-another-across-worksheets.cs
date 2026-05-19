using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class CopyShape3DFormat
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Source worksheet with the shape that has 3D formatting
        Worksheet srcSheet = workbook.Worksheets[0];
        srcSheet.Name = "Source";

        // Destination worksheet where the target shape resides
        Worksheet destSheet = workbook.Worksheets.Add("Destination");

        // Add a rectangle shape to the source sheet and configure its 3D format
        Shape srcShape = srcSheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 100);
        ThreeDFormat src3D = srcShape.ThreeDFormat;
        src3D.ExtrusionColor.Color = Color.Blue;
        src3D.LightAngle = 45;
        src3D.ContourWidth = 2;
        src3D.ContourColor.Color = Color.Red;
        src3D.Material = PresetMaterialType.Metal;
        src3D.ExtrusionHeight = 30;
        src3D.RotationX = 10;
        src3D.RotationY = 20;
        src3D.RotationZ = 30;
        src3D.BottomBevelHeight = 5;
        src3D.BottomBevelWidth = 5;
        src3D.BottomBevelType = BevelType.ArtDeco;
        src3D.TopBevelHeight = 8;
        src3D.TopBevelWidth = 8;
        src3D.TopBevelType = BevelType.SoftRound;
        src3D.Perspective = 30;
        src3D.PresetCameraType = PresetCameraType.PerspectiveFront;
        src3D.Z = 15;

        // Add a rectangle shape to the destination sheet (initially without 3D formatting)
        Shape destShape = destSheet.Shapes.AddRectangle(2, 0, 2, 150, 200, 100);
        ThreeDFormat dest3D = destShape.ThreeDFormat;

        // Copy all 3D format properties from the source shape to the destination shape
        dest3D.ExtrusionColor.Color = src3D.ExtrusionColor.Color;
        dest3D.LightAngle = src3D.LightAngle;
        dest3D.ContourWidth = src3D.ContourWidth;
        dest3D.ContourColor.Color = src3D.ContourColor.Color;
        dest3D.Material = src3D.Material;
        dest3D.ExtrusionHeight = src3D.ExtrusionHeight;
        dest3D.RotationX = src3D.RotationX;
        dest3D.RotationY = src3D.RotationY;
        dest3D.RotationZ = src3D.RotationZ;
        dest3D.BottomBevelHeight = src3D.BottomBevelHeight;
        dest3D.BottomBevelWidth = src3D.BottomBevelWidth;
        dest3D.BottomBevelType = src3D.BottomBevelType;
        dest3D.TopBevelHeight = src3D.TopBevelHeight;
        dest3D.TopBevelWidth = src3D.TopBevelWidth;
        dest3D.TopBevelType = src3D.TopBevelType;
        dest3D.Perspective = src3D.Perspective;
        dest3D.PresetCameraType = src3D.PresetCameraType;
        dest3D.Z = src3D.Z;

        // Save the workbook with the copied 3D format
        workbook.Save("Copy3DFormatDemo.xlsx");
    }
}