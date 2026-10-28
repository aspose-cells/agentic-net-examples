// Title: C# routine to toggle texture tiling of a named shape in an Excel workbook using Aspose.Cells
// Common Searches: C# Aspose.Cells toggle texture tiling for a shape by name | how to enable or disable shape texture tiling in an Excel file with Aspose.Cells | programmatically change the Tiling property of a shape's fill texture using Aspose.Cells .NET | Aspose.Cells example to flip texture tiling flag on a worksheet shape | toggle image texture tiling on Excel shape using Aspose.Cells API
// Tags: Aspose.Cells shape texture tiling toggle | C# modify shape fill texture property | Excel shape texture manipulation Aspose.Cells | Aspose.Cells workbook shape fill adjustment | toggle shape texture tiling .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The provided C# class loads an Excel workbook, finds a shape by its name, checks that the shape has a texture, flips the texture's Tiling flag, and saves the file, with detailed error handling for missing files, shapes, or textures.
public class ShapeTextureTiler
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <param name="shapeName">Name of the shape whose texture tiling will be toggled.</param>
    public static void ToggleTextureTiling(string filePath, string shapeName)
    {
        // Verify that the workbook file exists
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel file not found.", filePath);

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the shape by name
            Shape targetShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.Name.Equals(shapeName, StringComparison.OrdinalIgnoreCase))
                {
                    targetShape = shape;
                    break;
                }
            }

            if (targetShape == null)
                throw new ArgumentException($"Shape with name '{shapeName}' not found.");

            // Get the fill format of the shape
            FillFormat fill = targetShape.Fill;

            // Ensure the shape has a texture; otherwise we cannot toggle tiling
            if (fill.Texture == null)
                throw new InvalidOperationException("The specified shape does not have a texture to toggle.");

            // Use dynamic to avoid compile‑time dependency on the Texture class
            dynamic texture = fill.Texture;

            // Toggle the tiling flag (applicable for image textures)
            bool currentTiling = texture.Tiling;
            texture.Tiling = !currentTiling;

            // Save the workbook back to the same file
            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            // Wrap any exception with additional context
            throw new InvalidOperationException("Failed to toggle texture tiling.", ex);
        }
    }

    // Example usage
    public static void Main()
    {
        try
        {
            string excelPath = @"C:\Temp\Sample.xlsx";
            string shapeToToggle = "MyShape";

            ToggleTextureTiling(excelPath, shapeToToggle);
            Console.WriteLine("Texture tiling toggled successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
