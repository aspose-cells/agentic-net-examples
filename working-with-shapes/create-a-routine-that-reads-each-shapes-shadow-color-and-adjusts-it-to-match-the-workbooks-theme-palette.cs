// Title: Adjust Excel shape shadow colors to match a workbook theme palette using Aspose.Cells for .NET
// AI Prompts: Write a C# method that loops through every shape on all worksheets, reads each shape's Shadow.Color via reflection, finds the closest color from a predefined theme palette, and sets the shadow to that color. | Create a console application that accepts input and output .xlsx file paths, loads the workbook with Aspose.Cells, applies the shadow‑color‑matching routine, and includes robust file‑existence checks and exception handling.
// Common Searches: asp.net adjust shape shadow color based on workbook theme palette Aspose.Cells | c# use reflection to modify shape shadow property in Excel file | find nearest color from custom palette for Excel shape shadows using Aspose.Cells | batch update all shape shadows to theme colors in a .xlsx workbook with .NET
// Tags: shape shadow color normalization Aspose.Cells | reflection access shape shadow .NET | nearest palette color algorithm C# | batch process shapes across worksheets | excel workbook theme palette mapping

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example iterates through every shape in each worksheet, uses reflection to read and update the Shadow.Color property, selects the nearest color from a predefined palette, and saves the workbook with adjusted shape shadows.
public class ShapeShadowAdjuster
{
    // Adjusts each shape's shadow color to the nearest color in a predefined palette.
    // Uses only APIs available in all supported Aspose.Cells versions.
    public static void AdjustShapeShadows(string inputFilePath, string outputFilePath)
    {
        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFilePath))
            throw new FileNotFoundException($"Input file not found: {inputFilePath}");

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFilePath);

            // Define a simple palette of colors to approximate a theme palette
            List<Color> palette = new List<Color>
            {
                Color.Black,
                Color.White,
                Color.Red,
                Color.Green,
                Color.Blue,
                Color.Yellow,
                Color.Orange,
                Color.Purple,
                Color.Gray,
                Color.Brown
            };

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Aspose.Cells.Drawing.Shape shape in sheet.Shapes)
                {
                    // Attempt to adjust shadow if the shape supports a Shadow property (available in newer versions)
                    // Use reflection to avoid compile‑time dependency on the Shadow API.
                    var shadowProp = shape.GetType().GetProperty("Shadow");
                    if (shadowProp != null && shadowProp.CanRead && shadowProp.CanWrite)
                    {
                        var shadowObj = shadowProp.GetValue(shape);
                        if (shadowObj != null)
                        {
                            // Get current shadow color via reflection
                            var colorProp = shadowObj.GetType().GetProperty("Color");
                            if (colorProp != null && colorProp.CanRead && colorProp.CanWrite)
                            {
                                Color originalColor = (Color)colorProp.GetValue(shadowObj);

                                // Find nearest palette color (Euclidean distance in RGB space)
                                Color nearest = palette[0];
                                double minDist = double.MaxValue;
                                foreach (Color c in palette)
                                {
                                    double dist = Math.Sqrt(
                                        Math.Pow(originalColor.R - c.R, 2) +
                                        Math.Pow(originalColor.G - c.G, 2) +
                                        Math.Pow(originalColor.B - c.B, 2));
                                    if (dist < minDist)
                                    {
                                        minDist = dist;
                                        nearest = c;
                                    }
                                }

                                // Set the shadow color to the nearest palette color
                                colorProp.SetValue(shadowObj, nearest);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputFilePath);
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide context while preserving the original stack trace
            throw new ApplicationException("Error processing workbook for shape shadow adjustment.", ex);
        }
    }
}

public class Program
{
    // Entry point required for console application
    public static void Main(string[] args)
    {
        // Default file names; can be overridden via command‑line arguments
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        if (args.Length >= 2)
        {
            inputPath = args[0];
            outputPath = args[1];
        }

        try
        {
            ShapeShadowAdjuster.AdjustShapeShadows(inputPath, outputPath);
            Console.WriteLine($"Shape shadows adjusted successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to adjust shape shadows: {ex.Message}");
        }
    }
}
