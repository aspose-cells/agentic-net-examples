// Title: Strip all 3‑D formatting from Excel shapes while preserving size and position using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates every worksheet, and resets each shape's ThreeDFormat properties (Depth, RotationX/Y/Z, shadows, bevels, lighting, material) to their defaults without changing the shape's dimensions. | Show how to employ reflection in C# to safely assign values to optional ThreeDFormat members on a Shape object, ensuring all 3‑D effects are removed even when the API version varies. | Provide a reusable method that accepts a Workbook instance, removes all 3‑D effects from its shapes, and saves the modified file to a specified output path.
// Common Searches: C# Aspose.Cells remove threeDFormat from all shapes in a workbook | how to clear 3d effects on Excel shapes using Aspose.Cells .NET | reset shape depth rotation shadow bevel Aspose.Cells example | iterate through worksheets and shapes to disable 3d formatting Aspose.Cells | use reflection to set shape ThreeDFormat properties in Aspose.Cells
// Tags: Aspose.Cells clear shape 3d formatting | disable bevel and shadow on Excel shapes | iterate all worksheets shapes Aspose.Cells | reflection based property setting C# | strip threeDFormat from shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an input Excel file, walks through each worksheet and every shape, and uses reflection to set all ThreeDFormat members (Depth, rotations, shadow, contour, bevels, preset lighting, preset material, light rig) to their default values. This removes every 3‑D effect while leaving the shapes' 2‑D attributes such as size and position unchanged, then saves the cleaned workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    var threeD = shape.ThreeDFormat;
                    if (threeD == null) continue;

                    // Reset 3‑D formatting using reflection (handles API differences)
                    SetPropertyIfExists(threeD, "Depth", 0);
                    SetPropertyIfExists(threeD, "RotationX", 0);
                    SetPropertyIfExists(threeD, "RotationY", 0);
                    SetPropertyIfExists(threeD, "RotationZ", 0);
                    SetPropertyIfExists(threeD, "IsShadowEnabled", false);
                    SetPropertyIfExists(threeD, "IsContourEnabled", false);
                    SetPropertyIfExists(threeD, "IsBevelTopEnabled", false);
                    SetPropertyIfExists(threeD, "IsBevelBottomEnabled", false);
                    SetEnumPropertyIfExists(threeD, "PresetLighting", "Normal");
                    SetEnumPropertyIfExists(threeD, "PresetMaterial", "Plastic");
                    SetEnumPropertyIfExists(threeD, "LightRig", "ThreePt");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper to set simple properties via reflection
    static void SetPropertyIfExists(object obj, string propertyName, object value)
    {
        var prop = obj.GetType().GetProperty(propertyName);
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType));
        }
    }

    // Helper to set enum properties via reflection
    static void SetEnumPropertyIfExists(object obj, string propertyName, string enumValueName)
    {
        var prop = obj.GetType().GetProperty(propertyName);
        if (prop != null && prop.CanWrite && prop.PropertyType.IsEnum)
        {
            var enumValue = Enum.Parse(prop.PropertyType, enumValueName);
            prop.SetValue(obj, enumValue);
        }
    }
}
