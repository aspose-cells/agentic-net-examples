// Title: How to remove a shape's shadow while keeping its reflection and glow in Aspose.Cells for .NET
// AI Prompts: Set the IsVisible property of a shape's Shadow to false using the Aspose.Cells .NET API. | Use .NET reflection to disable a shape's shadow without altering its reflection or glow in an Excel workbook. | Programmatically turn off the shadow effect of the first shape in a worksheet while preserving other visual effects with Aspose.Cells.
// Common Searches: aspnet remove shadow from Excel shape but keep reflection | aspose.cells disable shape shadow while preserving glow | c# set shape shadow visibility false in Aspose.Cells | how to retain shape reflection when turning off shadow in Aspose.Cells
// Tags: disable shape shadow Aspose.Cells | shape shadow visibility .NET | preserve shape reflection Aspose.Cells | retain shape glow Aspose.Cells | reflection effect Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first shape on the first worksheet, uses reflection to locate the shape's Shadow object and sets its IsVisible property to false, thereby removing the shadow while leaving reflection and glow intact, and then saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                // Get the first shape
                Shape shape = worksheet.Shapes[0];

                // Attempt to disable shadow visibility using reflection (covers different API versions)
                try
                {
                    var shadowProp = shape.GetType().GetProperty("Shadow");
                    if (shadowProp != null)
                    {
                        var shadowObj = shadowProp.GetValue(shape);
                        if (shadowObj != null)
                        {
                            var isVisibleProp = shadowObj.GetType().GetProperty("IsVisible");
                            if (isVisibleProp != null && isVisibleProp.CanWrite)
                            {
                                isVisibleProp.SetValue(shadowObj, false);
                                Console.WriteLine("Shadow visibility disabled.");
                            }
                            else
                            {
                                Console.WriteLine("Shadow property exists but does not expose IsVisible.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Shadow object is null; cannot modify.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Shape does not support a Shadow property in this API version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to modify shadow effect: {ex.Message}");
                }

                // Save the modified workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
