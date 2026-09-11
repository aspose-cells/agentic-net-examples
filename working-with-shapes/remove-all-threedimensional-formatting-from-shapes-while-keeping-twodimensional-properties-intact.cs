// Title: Remove 3‑D formatting from every shape in an Excel workbook while keeping 2‑D rotation using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates through all worksheets and their Shape collections, sets each shape's ThreeDFormat.Material to Plastic, clears ThreeDFormat.ExtrusionColor, and resets RotationAngle to 0 before saving the workbook. | Show a complete Aspose.Cells example that strips all three‑dimensional effects from shapes in a workbook but leaves the existing 2‑D rotation unchanged.
// Common Searches: aspocells c# remove three dimensional effects from all worksheet shapes | how to clear shape extrusion color with Aspose.Cells .NET | reset shape rotation angle after removing 3d formatting using Aspose.Cells | iterate through shapes collection in Aspose.Cells to modify ThreeDFormat | remove 3d material from Excel shapes programmatically with Aspose.Cells
// Tags: Aspose.Cells remove 3D formatting from shapes | Aspose.Cells reset shape rotation angle | Aspose.Cells clear shape extrusion color | Aspose.Cells iterate worksheet shapes | Aspose.Cells workbook shape processing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel file, walks every worksheet's Shape collection, clears any ThreeDFormat by setting Material to Plastic and nulling ExtrusionColor, resets each shape's RotationAngle to zero, and saves the modified workbook.
class Remove3DFormatting
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Process each shape
                foreach (Shape shape in shapes)
                {
                    // Remove 3‑D formatting if available
                    if (shape.ThreeDFormat != null)
                    {
                        // Reset material to a default value
                        shape.ThreeDFormat.Material = PresetMaterialType.Plastic;

                        // Clear extrusion color (expects CellsColor)
                        shape.ThreeDFormat.ExtrusionColor = null;

                        // Note: LightRig property does not exist in the current API; omitted.
                    }

                    // Reset 2‑D rotation
                    shape.RotationAngle = 0;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
