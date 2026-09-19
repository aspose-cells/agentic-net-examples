// Title: Reset all shape visual effects to defaults in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write a C# routine that iterates through every worksheet and shape in an Aspose.Cells Workbook and restores each shape's EffectFormat properties (Type, Blur, Glow, Shadow, ThreeDFormat, Reflection, SoftEdges) to their default values. | Create a helper method that uses reflection to clear all visual effects from shapes in a workbook, handling different Aspose.Cells versions, and integrate it into a console application that loads, modifies, and saves an Excel file.
// Common Searches: Aspose.Cells C# reset shape effects for entire workbook | How to clear blur and shadow from all shapes in an Excel file using Aspose.Cells | C# code to remove 3D and reflection effects from shapes in every worksheet | Set shape EffectFormat.Type to None across all worksheets with Aspose.Cells | Utility method to reset visual effects of shapes in an Aspose.Cells workbook
// Tags: reset shape EffectFormat Aspose.Cells | clear blur glow shadow from Excel shapes .NET | iterate worksheets to reset shape properties C# | Aspose.Cells default shape visual effects | reflection soft edges removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, walks through each worksheet and its shapes, and via reflection resets the EffectFormat (type, blur, glow, shadow, 3D, reflection, soft edges) of every shape to default values before saving the file.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Reset shape effects across the workbook
            ResetAllShapeEffects(workbook);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Utility method to reset all shape effects in the entire workbook
    static void ResetAllShapeEffects(Workbook wb)
    {
        // Iterate through each worksheet
        foreach (Worksheet sheet in wb.Worksheets)
        {
            // Iterate through each shape on the worksheet
            foreach (Aspose.Cells.Drawing.Shape shape in sheet.Shapes)
            {
                // Use reflection to access EffectFormat if it exists (covers different library versions)
                var effectProp = shape.GetType().GetProperty("EffectFormat");
                if (effectProp != null && effectProp.CanRead && effectProp.CanWrite)
                {
                    var effect = effectProp.GetValue(shape);
                    if (effect != null)
                    {
                        // Reset the effect type to None
                        var typeProp = effect.GetType().GetProperty("Type");
                        var effectTypeEnum = effect.GetType().Assembly.GetType("Aspose.Cells.Drawing.EffectType");
                        if (typeProp != null && effectTypeEnum != null)
                        {
                            var noneValue = Enum.Parse(effectTypeEnum, "None");
                            typeProp.SetValue(effect, noneValue);
                        }

                        // Reset numeric and object properties to defaults if they exist
                        SetIfExists(effect, "Blur", 0);
                        SetIfExists(effect, "Glow", null);
                        SetIfExists(effect, "Shadow", null);
                        SetIfExists(effect, "ThreeDFormat", null);
                        SetIfExists(effect, "Reflection", null);
                        SetIfExists(effect, "SoftEdges", null);
                    }
                }
            }
        }
    }

    // Helper method to set a property via reflection if it exists
    static void SetIfExists(object obj, string propertyName, object value)
    {
        var prop = obj.GetType().GetProperty(propertyName);
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(obj, value);
        }
    }
}
