// Title: Apply a consistent shadow offset to every shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net set same shadow offset for all shapes in an Excel workbook using Aspose.Cells | how to batch apply shadow distance and angle to Excel shapes with Aspose.Cells | Aspose.Cells C# apply diagonal shadow preset to multiple shapes | programmatically change shape shadow properties in an .xlsx file using Aspose.Cells | uniform shape shadow formatting across worksheets with Aspose.Cells .NET
// Tags: apply consistent shadow offset to Excel shapes Aspose.Cells | batch set shape shadow properties C# | preset diagonal shadow effect Aspose.Cells | modify shape shadow distance and angle .NET | programmatic shape shadow formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ApplyUniformShadowOffset
{
    // The example loads (or creates) an Excel workbook, iterates through every worksheet and each shape, and sets each shape's ShadowEffect to a diagonal offset preset with a uniform distance of 12 points and a 45‑degree angle. Optional transparency and blur are applied, then the workbook is saved with the updated shadow formatting.
    class Program
    {
        // Uniform shadow distance in points
        private const double UniformShadowDistance = 12.0;
        // Uniform shadow angle in degrees (optional, keeps visual alignment)
        private const double UniformShadowAngle = 45.0;

        static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
                string outputPath = args.Length > 1 ? args[1] : "output_with_shadows.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one with sample shapes
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook with sample shapes.");
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];

                    // Add a rectangle shape
                    ws.Shapes.AddRectangle(1, 0, 1, 0, 150, 80);

                    // Add an oval shape (AutoShapeType.Oval replaces the non‑existent Ellipse)
                    ws.Shapes.AddAutoShape(AutoShapeType.Oval, 2, 0, 2, 0, 120, 120);

                    // Add a horizontal line shape
                    ws.Shapes.AddLine(3, 0, 3, 0, 0, 200);
                }

                // Apply uniform shadow offset to every shape in every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Processing worksheet \"{sheet.Name}\" with {sheet.Shapes.Count} shape(s).");

                    for (int i = 0; i < sheet.Shapes.Count; i++)
                    {
                        Shape shape = sheet.Shapes[i];

                        try
                        {
                            // Access the shadow effect of the shape
                            ShadowEffect shadow = shape.ShadowEffect;

                            // Use a preset that supports offset
                            shadow.PresetType = PresetShadowType.OffsetDiagonalBottomRight;

                            // Apply uniform distance and angle
                            shadow.Distance = UniformShadowDistance;
                            shadow.Angle = UniformShadowAngle;

                            // Optional visual tweaks
                            shadow.Transparency = 0.3; // semi‑transparent
                            shadow.Blur = 5;           // slight blur

                            Console.WriteLine($"  Shape {i}: Type={shape.Type}, Name=\"{shape.Name}\" – Shadow set (Distance={shadow.Distance}, Angle={shadow.Angle})");
                        }
                        catch (Exception shapeEx)
                        {
                            Console.Error.WriteLine($"Error applying shadow to shape {i} in worksheet \"{sheet.Name}\": {shapeEx.Message}");
                        }
                    }
                }

                // Save the modified workbook
                Console.WriteLine($"Saving workbook with updated shadows to \"{outputPath}\"...");
                workbook.Save(outputPath);
                Console.WriteLine("Operation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while applying shadow offsets:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
