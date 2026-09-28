// Title: Ungroup a named GroupShape in an Excel workbook using Aspose.Cells for .NET, reposition each resulting shape, and save the file
// AI Prompts: Generate C# code with Aspose.Cells that opens an Excel file, locates a GroupShape by its name on the first worksheet, calls Ungroup(), shifts every new shape 10 points right and 5 points down, and saves the workbook to a specified output path. | Implement a C# method that takes inputPath, outputPath, and groupShapeName, loads the workbook, ungroups the identified GroupShape, updates the Left and Top properties of each resulting shape, and writes the modified workbook.
// Common Searches: Aspose.Cells C# ungroup GroupShape by name and move shapes | C# example to ungroup Excel shape group and adjust individual shape positions with Aspose.Cells | How to change Left and Top of shapes after ungrouping using Aspose.Cells for .NET | Saving an Excel workbook after modifying shape coordinates with Aspose.Cells
// Tags: Aspose.Cells ungroup GroupShape .NET | adjust shape Left Top Aspose.Cells | modify individual shapes Excel Aspose.Cells | save workbook after shape manipulation Aspose.Cells | retrieve GroupShape by name Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an existing Excel workbook, finds a GroupShape by its name on the first worksheet, ungroups it, moves each newly created shape 10 points to the right and 5 points down, and saves the updated workbook to the designated output file.
class Program
{
    static void Main()
    {
        try
        {
            // Define file paths (replace placeholders with actual paths)
            string inputPath = "{InputFilePath}";
            string outputPath = "{OutputFilePath}";
            string groupShapeName = "{GroupShapeName}";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the group shape by its name
            GroupShape groupShape = sheet.Shapes[groupShapeName] as GroupShape;
            if (groupShape == null)
            {
                Console.WriteLine($"Group shape \"{groupShapeName}\" not found or is not a group.");
                return;
            }

            // Record the number of shapes before ungrouping
            int originalShapeCount = sheet.Shapes.Count;

            // Ungroup the shape (method returns void)
            try
            {
                groupShape.Ungroup();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to ungroup shape: {ex.Message}");
                return;
            }

            // Adjust positions of the newly created individual shapes
            for (int i = originalShapeCount; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];
                shape.Left += 10; // move 10 points to the right
                shape.Top += 5;   // move 5 points down
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
