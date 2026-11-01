// Title: How to apply a custom reflection preset (size 40 %, blur 3 pt, distance 6 pt) to a rectangle shape using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a rectangle shape, and set its Reflection.Size to 40, Reflection.Blur to 3, and Reflection.Distance to 6 with Aspose.Cells in C#. | Configure the reflection effect of a worksheet shape by assigning custom size, blur, and distance values in a .NET application using Aspose.Cells.
// Common Searches: Aspose.Cells C# set rectangle shape reflection size to 40 percent | how to change blur and distance of shape reflection in Excel with Aspose.Cells | C# code for applying custom reflection preset to a worksheet shape using Aspose.Cells | apply reflection effect to Excel shape programmatically Aspose.Cells .NET
// Tags: Aspose.Cells shape reflection properties | set reflection size blur distance Aspose.Cells | custom reflection preset Excel shape C# | apply reflection effect to worksheet shape .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, inserts a rectangle shape on the first worksheet, configures its reflection effect with a size of 40 %, blur of 3 points, and distance of 6 points, and saves the file as CustomReflection.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,   // upper-left row
                    1,   // upper-left column
                    0,   // upper-left row offset (in pixels)
                    0,   // upper-left column offset (in pixels)
                    200, // width (in pixels)
                    100  // height (in pixels)
                );

                // Access the reflection effect of the shape
                ReflectionEffect reflection = shape.Reflection;

                // Apply a custom reflection preset
                reflection.Size = 40;    // size as a percentage (0‑100)
                reflection.Blur = 3;     // blur amount in points
                reflection.Distance = 6; // distance from the object in points

                // Save the workbook to a file
                string outputPath = "CustomReflection.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
