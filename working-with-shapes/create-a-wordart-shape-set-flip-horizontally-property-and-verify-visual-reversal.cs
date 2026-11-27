// Title: Add WordArt to an Excel worksheet and simulate a horizontal flip by rotating 180° with Aspose.Cells for .NET
// Common Searches: how to rotate a WordArt shape 180 degrees using Aspose.Cells C# | Aspose.Cells C# simulate horizontal flip for WordArt | add WordArt with preset style and compare rotated version in Excel using Aspose.Cells | verify IsWordArt flag and rotation angle of a shape in Aspose.Cells .NET | save workbook containing rotated WordArt as XLSX with Aspose.Cells
// Tags: WordArt shape rotation Aspose.Cells | horizontal flip simulation for WordArt | preset WordArt style addition C# | shape IsWordArt verification | export workbook to XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WordArtFlipDemo
{
    // The example creates a new workbook, inserts a WordArt shape with a preset style, sets its RotationAngle to 180° to mimic a horizontal flip, adds a second unrotated WordArt for visual comparison, prints the IsWordArt flag and rotation angles, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "WordArtFlipDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Add a WordArt shape with a preset style
                Shape wordArt = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle2,
                    "Flip Me!",
                    2,   // top row index
                    10,  // top offset in pixels
                    2,   // left column index
                    10,  // left offset in pixels
                    100, // height in pixels
                    300  // width in pixels
                );

                // 3. Verify that the shape is a WordArt object
                Console.WriteLine($"IsWordArt: {wordArt.IsWordArt}");

                // 4. Aspose.Cells does not provide a direct FlipMode property.
                //    As an alternative, rotate the shape 180 degrees to simulate a horizontal flip.
                wordArt.RotationAngle = 180;
                Console.WriteLine($"RotationAngle after setting: {wordArt.RotationAngle}");

                // 5. Add a second WordArt without rotation for visual comparison
                Shape wordArtNormal = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle3,
                    "Normal",
                    6,   // top row index
                    10,  // top offset in pixels
                    2,   // left column index
                    10,  // left offset in pixels
                    100, // height in pixels
                    300  // width in pixels
                );

                // Ensure the normal shape has no rotation
                wordArtNormal.RotationAngle = 0;
                Console.WriteLine($"Normal shape RotationAngle: {wordArtNormal.RotationAngle}");

                // 6. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'. Open it in Excel to see the rotated WordArt.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the WordArt shape:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
