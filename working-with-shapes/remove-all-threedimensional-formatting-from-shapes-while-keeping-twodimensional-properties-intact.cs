// Title: Remove 3‑D formatting from Excel shapes while preserving 2‑D properties using Aspose.Cells for .NET (C#)
// Common Searches: C# code to clear 3D effects from shapes in an Excel file with Aspose.Cells | how to disable ThreeDFormat on worksheet shapes using Aspose.Cells .NET | remove depth and rotation from Excel shapes programmatically with Aspose.Cells | preserve shape size and position while stripping 3D formatting in C# | Aspose.Cells example to reset shape ThreeDFormat properties to default
// Tags: Aspose.Cells remove shape ThreeDFormat | reset shape 3D properties C# | clear depth rotation perspective Aspose.Cells | preserve 2D shape attributes Excel .NET | iterate worksheet shapes Aspose.Cells example

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Remove3DFromShapes
{
    // The sample loads an Excel workbook, iterates through each worksheet and its shapes, and uses reflection to set the ThreeDFormat properties (Depth, RotationX, RotationY, RotationZ, Perspective) to zero and the Is3D flag to false. This removes all 3‑D formatting while keeping the original 2‑D size, position, and other attributes, then saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: Remove3DFromShapes <input.xlsx> <output.xlsx>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape has a ThreeDFormat object
                        if (shape.ThreeDFormat != null)
                        {
                            // Reset numeric 3‑D related properties via reflection (compatible with various SDK versions)
                            string[] numericProps = { "Depth", "RotationX", "RotationY", "RotationZ", "Perspective" };
                            foreach (string propName in numericProps)
                            {
                                PropertyInfo prop = shape.ThreeDFormat.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                                if (prop != null && prop.CanWrite)
                                {
                                    // Set numeric properties to zero (use appropriate type)
                                    object zeroValue = Convert.ChangeType(0, prop.PropertyType);
                                    prop.SetValue(shape.ThreeDFormat, zeroValue);
                                }
                            }

                            // Set Is3D flag to false if the property exists
                            PropertyInfo is3DProp = shape.ThreeDFormat.GetType().GetProperty("Is3D", BindingFlags.Public | BindingFlags.Instance);
                            if (is3DProp != null && is3DProp.CanWrite)
                            {
                                is3DProp.SetValue(shape.ThreeDFormat, false);
                            }

                            Console.WriteLine($"3‑D formatting removed from shape \"{shape.Name}\" on sheet \"{sheet.Name}\".");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without 3‑D shape formatting to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
