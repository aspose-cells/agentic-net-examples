// Title: Disable shape reflection effects without affecting other formatting in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells, iterates all worksheets and shapes, disables the Reflection effect on each shape via the EffectFormat API, and saves the modified workbook. | Demonstrate how to use .NET reflection to safely access a shape's EffectFormat and its Reflection.IsEnabled property in Aspose.Cells, turning off reflection while leaving other effects intact. | Create a complete example that checks for the existence of EffectFormat and Reflection on each shape, disables the reflection, handles missing members gracefully, and writes the result to a new file.
// Common Searches: Aspose.Cells C# disable shape reflection while preserving other effects | remove reflection effect from Excel shapes using Aspose.Cells .NET | how to turn off shape reflection in a workbook with Aspose.Cells for C# | iterate through all shapes in an Excel file and disable reflection Aspose.Cells example
// Tags: disable shape reflection Aspose.Cells C# | shape EffectFormat Reflection Aspose.Cells | iterate worksheet shapes Aspose.Cells | preserve other shape effects while modifying reflection | load and save workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample loads an existing XLSX workbook, walks through every worksheet and each shape, uses .NET reflection to locate the EffectFormat and its Reflection object, sets Reflection.IsEnabled to false while leaving other visual effects unchanged, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in worksheet.Shapes)
                    {
                        try
                        {
                            // Use reflection to access EffectFormat and Reflection if they exist in the current API version
                            var effectFormatProp = shape.GetType().GetProperty("EffectFormat");
                            if (effectFormatProp == null) continue;

                            var effectFormat = effectFormatProp.GetValue(shape);
                            if (effectFormat == null) continue;

                            var reflectionProp = effectFormat.GetType().GetProperty("Reflection");
                            if (reflectionProp == null) continue;

                            var reflection = reflectionProp.GetValue(effectFormat);
                            if (reflection == null) continue;

                            var isEnabledProp = reflection.GetType().GetProperty("IsEnabled");
                            if (isEnabledProp != null && isEnabledProp.CanWrite)
                            {
                                // Disable the reflection effect while leaving other effects untouched
                                isEnabledProp.SetValue(reflection, false);
                            }
                        }
                        catch (Exception exShape)
                        {
                            // Log shape-specific errors but continue processing other shapes
                            Console.WriteLine($"Error processing shape '{shape.Name}': {exShape.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
