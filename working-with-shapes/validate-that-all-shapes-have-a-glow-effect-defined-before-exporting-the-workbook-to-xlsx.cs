// Title: Check that every shape in an Aspose.Cells workbook has a Glow effect with a positive size before saving as XLSX (C#)
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all worksheet shapes, uses reflection to confirm each shape has an EffectFormat.Glow defined, and throws an exception if the Glow size is zero or negative. | Create a C# routine that logs each shape's name and Glow size, validates the Glow effect via reflection, and saves the workbook to XLSX only when all shapes pass the check. | Write a C# method that adds a default Glow effect to any shape missing it, using Aspose.Cells and reflection, then exports the workbook to XLSX.
// Common Searches: aspnet c# verify shape glow effect exists in Aspose.Cells workbook before export | using reflection to access EffectFormat Glow property for shapes in Aspose.Cells | how to ensure all shapes have positive glow size in an Excel file with Aspose.Cells | throw exception if shape missing glow effect Aspose.Cells C# | validate shape effects prior to saving workbook as XLSX using Aspose.Cells
// Tags: Aspose.Cells shape glow check | C# reflection EffectFormat access | enforce glow size greater than zero | export workbook to XLSX after shape validation | shape effect verification in Excel

using Aspose.Cells;
using System;
using System.IO;

// C# example that uses Aspose.Cells and reflection to ensure every shape has an EffectFormat.Glow with a size > 0 before saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Validate that every shape has a glow effect defined using reflection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Aspose.Cells.Drawing.Shape shape in sheet.Shapes)
                {
                    // Retrieve EffectFormat via reflection (property may not exist in older versions)
                    var effectFormatProp = shape.GetType().GetProperty("EffectFormat");
                    if (effectFormatProp == null)
                        throw new InvalidOperationException(
                            $"Shape '{shape.Name}' on sheet '{sheet.Name}' does not support EffectFormat.");

                    var effectFormat = effectFormatProp.GetValue(shape);
                    if (effectFormat == null)
                        throw new InvalidOperationException(
                            $"Shape '{shape.Name}' on sheet '{sheet.Name}' does not have an EffectFormat instance.");

                    // Retrieve Glow via reflection
                    var glowProp = effectFormat.GetType().GetProperty("Glow");
                    if (glowProp == null)
                        throw new InvalidOperationException(
                            $"Shape '{shape.Name}' on sheet '{sheet.Name}' does not support Glow effect.");

                    var glow = glowProp.GetValue(effectFormat);
                    if (glow == null)
                        throw new InvalidOperationException(
                            $"Shape '{shape.Name}' on sheet '{sheet.Name}' does not have a Glow effect defined.");

                    // Retrieve Size via reflection and validate
                    var sizeProp = glow.GetType().GetProperty("Size");
                    if (sizeProp == null)
                        throw new InvalidOperationException(
                            $"Glow effect on shape '{shape.Name}' does not expose a Size property.");

                    var sizeValue = sizeProp.GetValue(glow);
                    double size = Convert.ToDouble(sizeValue);
                    if (size <= 0)
                        throw new InvalidOperationException(
                            $"Shape '{shape.Name}' on sheet '{sheet.Name}' has a glow effect with non‑positive size.");
                }
            }

            // Save the workbook after successful validation
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
