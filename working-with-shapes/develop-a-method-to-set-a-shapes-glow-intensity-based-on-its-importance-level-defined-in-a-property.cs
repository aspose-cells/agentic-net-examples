// Title: Set an Excel shape's glow intensity based on importance level using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that opens an Excel workbook with Aspose.Cells, locates a shape by name on a specified worksheet, and assigns a Glow.Size value that matches an importance string such as High, Medium, or Low. | Create a helper that converts importance levels to numeric glow sizes, updates the shape's Glow property, and saves the modified workbook to a new file while handling missing files and invalid worksheets.
// Common Searches: aspocells set shape glow intensity c# | c# adjust Excel shape glow based on custom importance level | how to change shape glow size in Aspose.Cells workbook | apply conditional glow effect to Excel shape using Aspose.Cells .NET | map importance level to shape visual effect Aspose.Cells example
// Tags: Aspose.Cells set shape glow size | C# map importance level to shape glow | Excel shape visual styling with Aspose.Cells | modify shape glow property .NET | conditional shape formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Provides a ShapeGlowHelper that maps importance strings (High, Medium, Low) to numeric glow sizes, loads a workbook, finds a named shape on a given worksheet, sets the shape's Glow.Size accordingly, ensures the output path exists, and saves the updated file, with error handling for missing resources.
public static class ShapeGlowHelper
{
    // Map importance level to glow size (intensity)
    private static float GetGlowSize(string importance)
    {
        if (string.IsNullOrEmpty(importance))
            return 0f;

        switch (importance.Trim().ToLower())
        {
            case "high":
                return 20f;   // strong glow
            case "medium":
                return 10f;   // moderate glow
            case "low":
                return 5f;    // subtle glow
            default:
                return 0f;    // no glow
        }
    }

    // Sets the glow intensity of a shape based on its importance level
    public static void SetShapeGlowByImportance(
        string workbookPath,   // path to the source workbook
        string sheetName,      // name of the worksheet containing the shape
        string shapeName,      // name of the shape to modify
        string importanceLevel,// importance level (e.g., "High", "Medium", "Low")
        string outputPath)     // path to save the modified workbook
    {
        try
        {
            // Verify source workbook exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook not found: {workbookPath}");

            // Load the workbook
            Workbook wb = new Workbook(workbookPath);

            // Get the target worksheet
            Worksheet ws = wb.Worksheets[sheetName];
            if (ws == null)
                throw new Exception($"Worksheet '{sheetName}' not found.");

            // Locate the shape by its name
            Shape targetShape = null;
            foreach (Shape shape in ws.Shapes)
            {
                if (shape.Name == shapeName)
                {
                    targetShape = shape;
                    break;
                }
            }

            if (targetShape == null)
                throw new Exception($"Shape '{shapeName}' not found in worksheet '{sheetName}'.");

            // Determine the glow size based on importance
            float glowSize = GetGlowSize(importanceLevel);

            // Apply or remove the glow effect
            // The Glow property is read‑only; modify its Size directly.
            if (targetShape.Glow != null)
            {
                targetShape.Glow.Size = glowSize;
            }

            // Ensure the output directory exists
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the modified workbook
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or rethrow as needed; for now we just write to console.
            Console.Error.WriteLine($"Error processing shape glow: {ex.Message}");
            throw;
        }
    }
}

// Minimal entry point required by the project.
public class Program
{
    public static void Main(string[] args)
    {
        // Example usage (adjust paths and names as needed).
        // This block can be removed or replaced by actual test code.
        try
        {
            string workbookPath = "Input.xlsx";
            string sheetName = "Sheet1";
            string shapeName = "MyShape";
            string importanceLevel = "High";
            string outputPath = "Output.xlsx";

            ShapeGlowHelper.SetShapeGlowByImportance(
                workbookPath,
                sheetName,
                shapeName,
                importanceLevel,
                outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
