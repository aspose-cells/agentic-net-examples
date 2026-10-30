// Title: Clone an Excel worksheet shape, apply a reflection effect, and offset the copy by two columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clone worksheet shape and set reflection properties | How to duplicate a shape and move it to a different column in Excel with Aspose.Cells | Applying reflection effect to a copied shape using Aspose.Cells for .NET | C# code to add a shape copy with transparency and blur in Aspose.Cells | Offset cloned shape by columns in an Excel workbook using Aspose.Cells API
// Tags: AddCopy method shape cloning Aspose.Cells | ReflectionEffect property configuration C# | Worksheet shape column offset positioning | Excel shape transparency and blur settings | Aspose.Cells shape duplication example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, checks for at least one shape, clones the first shape with the AddCopy method, configures its ReflectionEffect (transparency, size, distance, blur), moves the cloned shape two columns to the right of the original, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one shape on the worksheet
                if (worksheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found on the worksheet.");
                    return;
                }

                // Get the first shape as the original shape
                Shape originalShape = worksheet.Shapes[0];

                // Clone the original shape using AddCopy overload (offsets set to 0)
                Shape clonedShape = worksheet.Shapes.AddCopy(
                    originalShape,
                    originalShape.UpperLeftRow,
                    originalShape.UpperLeftColumn,
                    0,
                    0);

                // Apply reflection effect if supported
                ReflectionEffect reflection = clonedShape.Reflection;
                if (reflection != null)
                {
                    reflection.Transparency = 0.5; // 50% transparent
                    reflection.Size = 0.8;         // 80% size of the original shape
                    reflection.Distance = 0.2;    // Distance from the shape
                    reflection.Blur = 2.0;        // Blur radius
                }
                else
                {
                    Console.WriteLine("Reflection effect is not supported for this shape.");
                }

                // Position the cloned shape two columns to the right of the original shape
                clonedShape.UpperLeftRow = originalShape.UpperLeftRow;
                clonedShape.UpperLeftColumn = originalShape.UpperLeftColumn + 2;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
