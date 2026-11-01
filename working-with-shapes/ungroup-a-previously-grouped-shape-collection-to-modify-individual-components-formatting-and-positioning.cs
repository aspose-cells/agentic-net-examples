// Title: Ungroup a GroupShape and adjust individual shape positions and sizes in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# example to ungroup a GroupShape in an Excel worksheet using Aspose.Cells | How to modify Left, Top, Width, and Height of shapes after ungrouping with Aspose.Cells | Iterate over worksheet shapes and change their dimensions in an XLSX file using Aspose.Cells .NET | Aspose.Cells code to move all shapes 10 points right and down after ungroup | Ungroup shapes and apply formatting changes in a .NET Excel file with Aspose.Cells
// Tags: Ungroup GroupShape Aspose.Cells | Set shape Left property Aspose.Cells | Set shape Top property Aspose.Cells | Resize shape Width Aspose.Cells | Resize shape Height Aspose.Cells | Worksheet shapes iteration Aspose.Cells | Excel xlsx shape ungroup C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample loads an existing XLSX workbook, finds the first GroupShape on the first worksheet, calls Ungroup to convert its child shapes into regular shapes, then iterates through all worksheet shapes adjusting their Left, Top, Width, and Height properties before saving the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the first GroupShape in the worksheet
                GroupShape groupShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape is GroupShape)
                    {
                        groupShape = (GroupShape)shape;
                        break; // Assuming only one group needs to be ungrouped
                    }
                }

                // Ungroup the shape collection if a group was found
                if (groupShape != null)
                {
                    try
                    {
                        // Ungroup the group shape; the child shapes become regular shapes in the worksheet
                        groupShape.Ungroup();

                        // Apply positioning changes to all shapes (including those just ungrouped)
                        foreach (Shape shp in worksheet.Shapes)
                        {
                            // Example positioning changes
                            shp.Left += 10;   // Move shape 10 points to the right
                            shp.Top += 10;    // Move shape 10 points down
                            shp.Width += 5;   // Increase width by 5 points
                            shp.Height += 5;  // Increase height by 5 points
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error while ungrouping or modifying shapes: {ex.Message}");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
