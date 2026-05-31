using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
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
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    0,   // upper left row
                    0,   // upper left column
                    10,  // top (pixels)
                    10,  // left (pixels)
                    200, // width (pixels)
                    100  // height (pixels)
                );
                shape.Text = "3D Material & Lighting Demo";

                // Access and configure the 3‑D format of the shape
                ThreeDFormat threeD = shape.ThreeDFormat;
                threeD.Material = PresetMaterialType.Plastic;
                threeD.Lighting = LightRigType.Soft;
                threeD.LightAngle = 70;          // approximate intensity
                threeD.ExtrusionHeight = 20;
                threeD.TopBevelType = BevelType.SoftRound;
                threeD.TopBevelWidth = 10;
                threeD.TopBevelHeight = 10;

                // Save the workbook
                string outputPath = "ThreeDMaterialSoftLighting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}