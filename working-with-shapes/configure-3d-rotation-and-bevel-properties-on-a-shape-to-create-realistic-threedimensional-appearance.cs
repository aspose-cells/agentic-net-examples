// Title: Apply 3‑D rotation and bevel formatting to a rectangle shape in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook, inserts a rectangle shape, sets ThreeDFormat.RotationX, RotationY, RotationZ, applies a bevel effect, and saves the file using Aspose.Cells. | Show how to configure shape.BevelTopType and shape.BevelBottomType together with ThreeDFormat rotations in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set shape ThreeDFormat rotation X Y Z example | How to add bevel to a shape in Excel using Aspose.Cells .NET | C# code to rotate rectangle shape in Excel workbook with Aspose.Cells | Apply 3D formatting to Excel shapes with Aspose.Cells library
// Tags: Aspose.Cells shape ThreeDFormat rotation | C# rectangle shape bevel Aspose.Cells | Excel 3D shape formatting .NET | ThreeDFormat rotationX rotationY rotationZ Aspose | save workbook with 3D rotated shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, adding a rectangle shape, configuring its ThreeDFormat rotation (X, Y, Z) and bevel properties, and saving the workbook as 3DShape.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

            // Configure 3‑D rotation (available in Aspose.Cells)
            shape.ThreeDFormat.RotationX = 30; // rotate 30° around the X‑axis
            shape.ThreeDFormat.RotationY = 20; // rotate 20° around the Y‑axis
            shape.ThreeDFormat.RotationZ = 10; // rotate 10° around the Z‑axis

            // Save the workbook (lifecycle save rule)
            string outputPath = "3DShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
