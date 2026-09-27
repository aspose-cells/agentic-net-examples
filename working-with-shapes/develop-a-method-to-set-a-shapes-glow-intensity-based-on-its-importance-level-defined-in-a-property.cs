// Title: Set an orange glow on an Aspose.Cells shape using the shape's AlternativeText importance level in C#
// AI Prompts: Create a C# method that reads an integer from a shape's AlternativeText, converts it to an orange shade, and applies the color to the shape's FillFormat.ForeColor and LineFormat.ForeColor while increasing the line weight using Aspose.Cells. | Extend the glow helper to accept custom start and end colors for a gradient and support importance values from 0 to 10, updating both fill and line colors accordingly. | Write example code that adds several rectangle shapes to a worksheet, assigns different importance levels via AlternativeText, and calls the glow helper for each shape to display varying glow intensities.
// Common Searches: how to change shape fill color based on alternative text in Aspose.Cells C# | apply gradient glow to Excel shapes using Aspose.Cells library | set line weight and color of a rectangle shape programmatically with Aspose.Cells | use importance level stored in shape properties to style shapes in a workbook
// Tags: shape fill color from importance property Aspose.Cells | orange glow intensity based on importance level C# | line format weight setting Aspose.Cells | visual emphasis for Excel shapes using Aspose.Cells | custom gradient glow helper Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example defines a ShapeGlowHelper.ApplyEmphasisBasedOnImportance method that reads an integer importance level from a shape's AlternativeText, clamps it to 0‑5, computes an orange color whose green component decreases with higher importance, and applies that color to the shape's FillFormat and LineFormat while increasing the line weight. The program creates a workbook, adds a rectangle shape, stores an importance value of "3" in AlternativeText, invokes the helper, and saves the file as ShapeWithGlow.xlsx.
public class ShapeGlowHelper
{
    // Sets visual emphasis of a shape based on its importance level.
    // The importance level is stored in the shape's AlternativeText property as an integer.
    public static void ApplyEmphasisBasedOnImportance(Shape shape)
    {
        if (shape == null) throw new ArgumentNullException(nameof(shape));

        // Retrieve importance level (default to 0 if parsing fails)
        int importance = 0;
        if (!int.TryParse(shape.AlternativeText, out importance))
            importance = 0;

        // Clamp importance to the range 0‑5
        importance = Math.Max(0, Math.Min(5, importance));

        // Determine a color that becomes more intense with higher importance
        // From light orange (low) to deep orange (high)
        int red = 255;
        int green = Math.Max(0, 200 - (importance * 30));
        int blue = 0;
        Color emphasisColor = Color.FromArgb(red, green, blue);

        // Apply the color to the shape's fill and line to simulate a "glow" effect
        shape.FillFormat.ForeColor = emphasisColor;
        shape.LineFormat.ForeColor = emphasisColor;
        shape.LineFormat.Weight = 2.0; // slightly thicker border for visibility
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 150);

            // Store the importance level in the shape's AlternativeText property (e.g., 3)
            shape.AlternativeText = "3";

            // Apply visual emphasis based on the stored importance level
            ShapeGlowHelper.ApplyEmphasisBasedOnImportance(shape);

            // Save the workbook
            string outputPath = "ShapeWithGlow.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
