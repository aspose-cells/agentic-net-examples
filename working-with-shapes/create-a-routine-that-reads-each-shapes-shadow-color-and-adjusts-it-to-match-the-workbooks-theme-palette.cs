// Title: C# – Adjust Shape Shadow Colors to Workbook Theme Using Aspose.Cells
// Description: Loads an Excel file, walks through every worksheet and shape, reads each shape's shadow color, replaces it with the nearest color from the workbook's theme palette via Workbook.GetMatchingColor, and saves the updated workbook.
// Keywords: Aspose.Cells | C# shape shadow | theme palette | GetMatchingColor | CellsColor | Excel shape formatting | shadow color adjustment
// Common Searches: Aspose.Cells change shape shadow to theme color | C# set shape shadow from workbook palette | match Excel shape shadow with theme colors | GetMatchingColor example for shadows | adjust all shape shadows in a workbook
// Developer Intent: Replace every shape's shadow color with the closest theme palette entry.
// Use Cases: Ensure brand‑consistent shadow hues across all graphics in a corporate report. | Prepare legacy workbooks for distribution by aligning visual effects with the current Excel theme. | Automate cleanup of imported spreadsheets where custom shadow colors no longer match the document's theme.
// AI Prompts: Generate a C# method that iterates through all shapes in a workbook and sets each shadow to the nearest theme color using Aspose.Cells. | Explain the purpose of Workbook.GetMatchingColor and show how to apply its result to a shape's ShadowEffect. | Create error‑handling code for shapes lacking a shadow effect or having null colors while updating shadow colors.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel file, walks through every worksheet and shape, reads each shape's shadow color, replaces it with the nearest color from the workbook's theme palette via Workbook.GetMatchingColor, and saves the updated workbook.
public class ShapeShadowThemeAdjuster
{
    // Adjusts each shape's shadow color to the closest color in the workbook's theme palette.
    public static void AdjustShapeShadowColors(string inputFilePath, string outputFilePath)
    {
        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputFilePath);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Access the shape's shadow effect
                ShadowEffect shadow = shape.ShadowEffect;

                // Ensure the shadow effect and its color are available
                if (shadow != null && shadow.Color != null)
                {
                    // Get the current shadow color (System.Drawing.Color)
                    Color currentColor = shadow.Color.Color;

                    // Find the best matching color in the workbook's palette/theme
                    Color matchedColor = workbook.GetMatchingColor(currentColor);

                    // Create a new CellsColor instance (lifecycle rule: create)
                    CellsColor cellsColor = workbook.CreateCellsColor();

                    // Mark it as a shape color to ensure correct handling
                    cellsColor.IsShapeColor = true;

                    // Assign the matched color
                    cellsColor.Color = matchedColor;

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
