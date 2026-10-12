// Title: Disable shape reflection effects while preserving other visual effects in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates over every worksheet and each Shape, checks for a Reflection effect, disables the reflection without changing shadows, glows, or other effects, and then saves the workbook. | Show how to use .NET reflection to access the Shape.Reflection.IsEnabled property in Aspose.Cells and set it to false for all shapes, ensuring the solution works across different library versions.
// Common Searches: Aspose.Cells C# disable shape reflection while preserving shadow effect | remove reflection from all shapes in an Excel file using Aspose.Cells .NET | how to turn off shape reflection in a workbook programmatically Aspose.Cells | iterate through worksheet shapes and modify reflection property Aspose.Cells C# | set shape reflection IsEnabled false with Aspose.Cells without affecting other effects
// Tags: shape reflection removal Aspose.Cells | preserve shape shadow Aspose.Cells | iterate worksheet shapes C# | set reflection IsEnabled false Aspose.Cells | modify shape visual effects Excel .NET

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and its shapes, detects any Reflection effect, and uses .NET reflection to set the Reflection.IsEnabled property to false while leaving other effects such as shadow and glow untouched, then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Disable reflection effect if it exists (using reflection to stay compatible with different API versions)
                    if (shape.Reflection != null)
                    {
                        PropertyInfo isEnabledProp = shape.Reflection.GetType().GetProperty("IsEnabled", BindingFlags.Public | BindingFlags.Instance);
                        if (isEnabledProp != null && isEnabledProp.CanWrite)
                        {
                            isEnabledProp.SetValue(shape.Reflection, false);
                        }
                    }
                    // Other effects (Shadow, Glow, etc.) are left unchanged
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
