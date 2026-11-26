// Title: How to apply the PerspectiveDiagonalBottomLeft shadow preset to a shape using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set shape shadow to PerspectiveDiagonalBottomLeft preset | change Excel shape shadow preset to PerspectiveDiagonalBottomLeft with Aspose.Cells | C# example for using PresetShadowType.PerspectiveDiagonalBottomLeft in Aspose.Cells | compare default shadow and PerspectiveDiagonalBottomLeft shadow on Excel shapes using Aspose.Cells | apply custom shadow effects to rectangle shapes in an Aspose.Cells workbook
// Tags: Aspose.Cells PresetShadowType API | C# rectangle shape shadow properties | Diagonal bottom-left perspective shadow configuration | Excel shape visual effect comparison using Aspose.Cells | Workbook shape shadow customization .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShadowPresetDemo
{
    // The sample creates a new workbook, adds two rectangle shapes to the first worksheet, applies the default OffsetDiagonalTopRight shadow to the first shape and switches the second shape's shadow preset to PerspectiveDiagonalBottomLeft. It configures shadow color, transparency, blur, and distance, labels each shape, saves the workbook as an XLSX file, and guides the user to open the file and observe the visual difference between the two shadow presets.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Shape 1 – default shadow (no preset)
                // ------------------------------------------------------------
                Shape defaultShadowShape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 80);
                defaultShadowShape.Name = "DefaultShadow";
                defaultShadowShape.Placement = PlacementType.FreeFloating;

                // Apply a simple shadow effect to make the difference visible
                ShadowEffect defaultShadow = defaultShadowShape.ShadowEffect;
                defaultShadow.PresetType = PresetShadowType.OffsetDiagonalTopRight;
                CellsColor defaultColor = workbook.CreateCellsColor();
                defaultColor.Color = Color.Black; // automatic-like color
                defaultShadow.Color = defaultColor;
                defaultShadow.Transparency = 0.3;
                defaultShadow.Blur = 10;
                defaultShadow.Distance = 5;

                // ------------------------------------------------------------
                // Shape 2 – PerspectiveDiagonalLowerLeft preset
                // ------------------------------------------------------------
                Shape perspectiveShape = sheet.Shapes.AddRectangle(5, 0, 5, 0, 150, 80);
                perspectiveShape.Name = "PerspectiveDiagonalLowerLeft";
                perspectiveShape.Placement = PlacementType.FreeFloating;

                // Set the shadow preset to PerspectiveDiagonalLowerLeft
                ShadowEffect perspectiveShadow = perspectiveShape.ShadowEffect;
                perspectiveShadow.PresetType = PresetShadowType.PerspectiveDiagonalLowerLeft;
                CellsColor perspectiveColor = workbook.CreateCellsColor();
                perspectiveColor.Color = Color.Black;
                perspectiveShadow.Color = perspectiveColor;
                perspectiveShadow.Transparency = 0.3;
                perspectiveShadow.Blur = 10;
                perspectiveShadow.Distance = 5;

                // Add some text to each shape so the result is easier to interpret
                defaultShadowShape.Text = "Default Shadow";
                perspectiveShape.Text = "PerspectiveDiagonalLowerLeft";

                // Save the workbook
                string outputPath = "ShadowPresetPerspectiveDiagonalLowerLeft.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("Open the file and compare the two rectangles:");
                Console.WriteLine("- The left rectangle uses the default OffsetDiagonalTopRight shadow.");
                Console.WriteLine("- The right rectangle uses the PerspectiveDiagonalLowerLeft preset shadow.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
