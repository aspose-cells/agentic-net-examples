// Title: Apply the BoldWave preset style to every WordArt shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, loops through all worksheets and shapes, detects WordArt objects, and assigns the BoldWave preset to their TextEffect property, handling missing PresetStyle via reflection. | Create a method that receives a workbook path, updates all WordArt shapes to use the BoldWave text effect, and saves the modified file, including error handling for older Aspose.Cells versions. | Write a script that validates the source Excel file, iterates over each shape, applies the BoldWave preset style to WordArt, and writes the result to a new workbook, logging any shapes that could not be updated.
// Common Searches: how to set BoldWave text effect on WordArt shapes with Aspose.Cells C# | Aspose.Cells apply preset style to all WordArt objects in a workbook | C# iterate Excel shapes and change WordArt text effect using reflection | update WordArt preset style in older Aspose.Cells versions | apply custom text effect to WordArt in Excel via Aspose.Cells .NET
// Tags: apply preset text effect to WordArt Aspose.Cells | iterate worksheet shapes using Aspose.Cells C# | reflection technique for TextEffect property | bulk update WordArt styling in Excel .NET | handle missing PresetStyle in older Aspose.Cells versions

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, walks through each worksheet and its shapes, identifies WordArt objects, and uses reflection to set their TextEffect PresetStyle to the BoldWave preset when available, then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets and their shapes.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only WordArt shapes.
                    if (shape.IsWordArt)
                    {
                        try
                        {
                            // Attempt to set the BoldWave preset style via reflection
                            // (covers scenarios where the PresetStyle property may not exist in older versions).
                            object textEffect = shape.TextEffect;
                            PropertyInfo presetProp = textEffect.GetType().GetProperty("PresetStyle", BindingFlags.Public | BindingFlags.Instance);
                            if (presetProp != null && presetProp.CanWrite)
                            {
                                Type enumType = presetProp.PropertyType;
                                object enumValue = Enum.Parse(enumType, "BoldWave", ignoreCase: true);
                                presetProp.SetValue(textEffect, enumValue);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to apply preset style to shape on sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
