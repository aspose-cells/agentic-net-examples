// Title: Reset all shape effects (shadow, glow, reflection) to default values in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# reset shadow, glow and reflection for every shape in a workbook | remove all shape formatting effects from Excel worksheets with Aspose.Cells .NET | reset shape line and fill transparency to defaults using Aspose.Cells API | clear 3D shadow property of shapes across multiple worksheets in C# | utility method to restore default visual effects of shapes in an Excel file with Aspose.Cells
// Tags: Aspose.Cells reset shape effects | clear shape formatting Excel .NET | default shape visual effects | iterate workbook shape collections C# | remove 3d shadow Aspose.Cells shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ResetShapeEffectsDemo
{
    // The example loads an Excel workbook, iterates through each worksheet's ShapeCollection, and for every shape resets shadow, glow, reflection, line and fill transparency to their default values, clears any 3‑D shadow flag, then saves the workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: ResetShapeEffectsDemo <inputWorkbookPath> <outputWorkbookPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Reset shape effects in every worksheet
                ResetAllShapeEffects(workbook);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="workbook">The workbook whose shapes will be processed.</param>
        private static void ResetAllShapeEffects(Workbook workbook)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];
                    try
                    {
                        ResetShapeEffects(shape);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to reset effects for shape \"{shape.Name}\": {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Resets the visual effect properties of a single shape to their default states.
        /// </summary>
        /// <param name="shape">The shape to reset.</param>
        private static void ResetShapeEffects(Shape shape)
        {
            // Reset ShadowEffect
            ShadowEffect shadow = shape.ShadowEffect;
            // Use the enum value that represents no preset shadow (if available)
            shadow.PresetType = PresetShadowType.NoShadow;
            shadow.Angle = 0;
            shadow.Blur = 0;
            shadow.Distance = 0;
            shadow.Size = 0;
            shadow.Transparency = 0;
            shadow.Color = null;

            // Reset Glow effect
            GlowEffect glow = shape.Glow;
            glow.Size = 0;
            glow.Transparency = 0;
            glow.Color = null;

            // Reset Reflection effect
            ReflectionEffect reflection = shape.Reflection;
            // Use the enum value that represents no reflection (if available)
            reflection.Type = ReflectionEffectType.None;
            reflection.Transparency = 0;
            reflection.Size = 0;
            reflection.Blur = 0;
            reflection.Distance = 0;

            // Reset line format transparency (new API)
            if (shape.Line != null)
            {
                shape.Line.Transparency = 0;
            }

            // Reset fill format transparency (new API)
            if (shape.Fill != null)
            {
                shape.Fill.Transparency = 0;
            }

            // Clear any 3‑D shadow flags that may exist on specific shape types
            try
            {
                dynamic dynShape = shape;
                if (HasProperty(dynShape, "Shadow"))
                {
                    dynShape.Shadow = false;
                }
            }
            catch
            {
                // Ignored – not all shapes expose a boolean Shadow property.
            }
        }

        /// <summary>
        /// Helper that checks at runtime whether an object has a given property.
        /// </summary>
        private static bool HasProperty(object obj, string propertyName)
        {
            return obj.GetType().GetProperty(propertyName) != null;
        }
    }
}
