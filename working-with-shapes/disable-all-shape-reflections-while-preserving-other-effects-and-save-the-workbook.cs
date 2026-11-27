// Title: How to disable shape reflections while keeping other effects in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove all shape reflections from an existing Excel file | disable reflection effect on shapes in a workbook with Aspose.Cells .NET | iterate through worksheets and shapes to turn off reflections using Aspose.Cells API | save Excel workbook without shape reflections while preserving fill and line styles in C# | set Shape.Reflection.Type to None for every shape in Aspose.Cells
// Tags: disable shape reflections Aspose.Cells | shape reflection properties Aspose.Cells C# | iterate worksheet shapes Aspose.Cells | preserve shape formatting while removing reflections | save workbook without reflections Aspose.Cells | reflection effect type None Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace DisableShapeReflections
{
    // The example loads an Excel workbook (or creates a sample one), iterates through all worksheets and their shapes, sets each shape's Reflection.Type to None and resets related reflection parameters, leaving other formatting such as fill and line styles unchanged, and then saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output_no_reflection.xlsx";

            Workbook workbook = null;

            // Load existing workbook or create a sample one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook with shapes.");
                try
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Sample";

                    // Add a rectangle shape
                    Shape rect = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);
                    rect.Name = "Rectangle1";
                    rect.Fill.FillType = FillType.Solid;
                    rect.Fill.SolidFill.Color = Color.LightBlue;
                    rect.Reflection.Type = ReflectionEffectType.HalfReflectionTouching;

                    // Add an oval shape (ellipse)
                    Shape oval = sheet.Shapes.AddShape(MsoDrawingType.Oval, 2, 0, 2, 0, 120, 80);
                    oval.Name = "Oval1";
                    oval.Fill.FillType = FillType.Solid;
                    oval.Fill.SolidFill.Color = Color.LightCoral;
                    oval.Reflection.Type = ReflectionEffectType.FullReflection4PtOffset;

                    // Add a line shape
                    Shape line = sheet.Shapes.AddShape(MsoDrawingType.Line, 3, 0, 3, 0, 200, 0);
                    line.Name = "Line1";
                    line.Line.Weight = 2.0;
                    line.Reflection.Type = ReflectionEffectType.TightReflectionTouching;

                    // Save the sample workbook
                    workbook.Save(inputPath);
                    Console.WriteLine($"Sample workbook saved as \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                    return;
                }
            }

            if (workbook == null)
            {
                Console.WriteLine("Workbook is not initialized.");
                return;
            }

            // Disable reflections on every shape in every worksheet
            int totalShapes = 0;
            int modifiedShapes = 0;

            try
            {
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Shape shape in ws.Shapes)
                    {
                        totalShapes++;

                        // Disable only if a reflection is set
                        if (shape.Reflection != null && shape.Reflection.Type != ReflectionEffectType.None)
                        {
                            shape.Reflection.Type = ReflectionEffectType.None;
                            shape.Reflection.Blur = 0;
                            shape.Reflection.Distance = 0;
                            shape.Reflection.FadeDirection = 0;
                            shape.Reflection.Size = 0;
                            shape.Reflection.Transparency = 0;
                            shape.Reflection.RotWithShape = false;

                            modifiedShapes++;
                            Console.WriteLine($"Disabled reflection for shape \"{shape.Name}\" on worksheet \"{ws.Name}\".");
                        }
                    }
                }

                Console.WriteLine($"Processed {totalShapes} shape(s); disabled reflection on {modifiedShapes} shape(s).");

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without reflections as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing shapes: {ex.Message}");
            }
        }
    }
}
