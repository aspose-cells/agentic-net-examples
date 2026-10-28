// Title: Ungroup a GroupShape in an Excel workbook and reposition each individual shape using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to ungroup an Excel shape group and move each shape | How to extract shapes from a GroupShape and change their Left and Top values with Aspose.Cells | Ungroup shapes in a worksheet and adjust positions before saving using Aspose.Cells .NET | Aspose.Cells C# code to ungroup MyGroup and shift shapes horizontally and vertically | Programmatically modify shape coordinates after ungrouping in an Excel file with Aspose.Cells
// Tags: ungroup groupshape aspocells | adjust shape left top aspocells | worksheet shapes collection aspocells | save workbook after shape changes aspocells | c# aspocells shape handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, finds a GroupShape named "MyGroup", ungroups it to create individual shapes, offsets each new shape's Left and Top properties, and saves the modified workbook.
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Find the group shape by name
                GroupShape? targetGroup = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape is GroupShape group && shape.Name == "MyGroup") // replace "MyGroup" with your group's name
                    {
                        targetGroup = group;
                        break;
                    }
                }

                if (targetGroup != null)
                {
                    // Record current shape count to identify newly created shapes after ungrouping
                    int originalCount = worksheet.Shapes.Count;

                    try
                    {
                        // Ungroup the shape group; this adds individual shapes to the worksheet
                        targetGroup.Ungroup();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to ungroup shape: {ex.Message}");
                        return;
                    }

                    // Adjust positions of the newly created individual shapes
                    int newCount = worksheet.Shapes.Count;
                    for (int i = originalCount; i < newCount; i++)
                    {
                        Shape individual = worksheet.Shapes[i];
                        individual.Left += 10; // shift horizontally
                        individual.Top += 5;   // shift vertically
                    }
                }
                else
                {
                    Console.WriteLine("Shape group named 'MyGroup' was not found.");
                }

                // Ensure output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
