// Title: C# – Set Shape Reflection Distance to 12 Points and Verify Persistence with Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, applies a half‑reflection effect, sets the reflection distance to 12 points, saves the file, reloads it, and reads back the distance to confirm the setting is persisted.
// Keywords: Aspose.Cells shape reflection distance | C# set reflection distance 12 points | verify shape reflection property after save | Aspose.Cells automated visual test | reflection effect Aspose.Cells .NET
// Common Searches: how to set reflection distance for a shape in Aspose.Cells C# | read back shape reflection distance after workbook save | Aspose.Cells screenshot comparison for shape effects | C# Aspose.Cells reflection effect properties
// Developer Intent: Set a shape’s reflection distance to 12 points, save the workbook, and programmatically confirm that the value is retained.
// Use Cases: Apply an exact 12‑point reflection distance to a rectangle for consistent styling in generated spreadsheets. | Persist reflection settings across saves so downstream processes render the same visual effect. | Include the distance check in automated UI tests that compare rendered screenshots before and after changes.
// AI Prompts: Generate C# code that adds a rectangle shape, configures a half‑reflection effect, sets the distance to 12 points, saves the workbook, and validates the saved value. | Provide a method to load a saved workbook and assert that the rectangle’s reflection distance equals 12 points. | Explain how to configure Aspose.Cells rendering options to produce reliable screenshots for shapes with reflection effects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, applies a half‑reflection effect, sets the reflection distance to 12 points, saves the file, reloads it, and reads back the distance to confirm the setting is persisted.
class ReflectionDistanceDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape that will show the reflection effect
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 200, 100);

            // Configure the reflection effect with visible properties
            ReflectionEffect reflection = shape.Reflection;
            reflection.Type = ReflectionEffectType.HalfReflectionTouching;
            reflection.Transparency = 0.5;
            reflection.Size = 80;
            reflection.Blur = 2;

            // Set the reflection distance to twelve points
            reflection.Distance = 12;

            // Save the workbook
            string filePath = "ReflectionDistanceDemo.xlsx";
            workbook.Save(filePath);

            // Verify that the workbook was saved and reload it to check the persisted property
            if (File.Exists(filePath))
            {
                Workbook loadedWorkbook = new Workbook(filePath);
                Shape loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];
                double loadedDistance = loadedShape.Reflection.Distance;
                Console.WriteLine("Loaded reflection distance: " + loadedDistance);
            }
            else
            {
                Console.WriteLine("Failed to locate the saved workbook file.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
