// Title: How to ungroup an IconSet GroupShape and edit its individual icons with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# ungroup a GroupShape by its name | access individual icons after ungrouping an IconSet group shape in Aspose.Cells | change line weight of a shape after ungrouping with Aspose.Cells .NET | remove IconSet group shape and edit each icon using Aspose.Cells C# | retrieve and modify shapes from a grouped IconSet in a workbook with Aspose.Cells
// Tags: ungroup groupshape Aspose.Cells C# | edit individual icons Aspose.Cells | shape line weight modification Aspose.Cells | find shape by name Aspose.Cells | iconset groupshape handling Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads a workbook, locates a GroupShape named "IconSetGroup" on the first worksheet, calls Ungroup to break it into separate top‑level shapes, then modifies the line weight of the first resulting shape before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string groupShapeName = "IconSetGroup";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet == null)
                {
                    Console.WriteLine("Error: Worksheet not found.");
                    return;
                }

                // Find the index of the group shape by its name
                int groupShapeIndex = -1;
                for (int i = 0; i < sheet.Shapes.Count; i++)
                {
                    if (sheet.Shapes[i].Name.Equals(groupShapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        groupShapeIndex = i;
                        break;
                    }
                }

                if (groupShapeIndex != -1)
                {
                    // Cast the shape to GroupShape
                    GroupShape iconSetGroup = sheet.Shapes[groupShapeIndex] as GroupShape;
                    if (iconSetGroup != null)
                    {
                        // Ungroup the shape – after this call the individual icons become top‑level shapes
                        iconSetGroup.Ungroup();

                        // Example: modify the first resulting shape (e.g., change its line weight)
                        if (sheet.Shapes.Count > 0)
                        {
                            Shape firstIcon = sheet.Shapes[0];
                            firstIcon.Line.Weight = 2.0; // Set line thickness
                            // Uncomment if supported: firstIcon.Line.Color = Color.Red;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"The shape named \"{groupShapeName}\" is not a GroupShape.");
                    }
                }
                else
                {
                    Console.WriteLine($"Group shape \"{groupShapeName}\" was not found in the worksheet.");
                }

                // Save the workbook with the changes
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
