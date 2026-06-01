using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

public class ShapeShadowThemeAdjuster
{
    // Adjusts each shape's shadow color to the nearest color in the workbook's theme palette.
    public static void AdjustShapeShadowColors(string inputFilePath, string outputFilePath)
    {
        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputFilePath);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Access the shape's shadow effect
                ShadowEffect shadow = shape.ShadowEffect;

                // Ensure the shadow has a color defined
                if (shadow != null && shadow.Color != null)
                {
                    // Get the current shadow color (System.Drawing.Color)
                    Color currentColor = shadow.Color.Color;

                    // Find the closest matching color in the workbook's theme palette
                    Color matchedColor = workbook.GetMatchingColor(currentColor);

                    // Create a CellsColor instance and assign the matched color
                    CellsColor cellsColor = workbook.CreateCellsColor();
                    cellsColor.IsShapeColor = true;          // Treat as shape color
                    cellsColor.Color = matchedColor;         // Set the matched color

                    // Apply the new color to the shadow effect
                    shadow.Color = cellsColor;
                }
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save(outputFilePath);
    }

    // Example usage
    public static void Main()
    {
        string inputPath = "InputWorkbook.xlsx";
        string outputPath = "AdjustedShadowWorkbook.xlsx";

        AdjustShapeShadowColors(inputPath, outputPath);

        Console.WriteLine("Shadow colors adjusted and workbook saved to: " + outputPath);
    }
}