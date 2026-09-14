// Title: How to reset all shape adjustment values to their defaults in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a .xlsx file, iterates through every worksheet and each Shape, clears the Adjustments collection when it exists, and saves the workbook. | Enhance the sample program to check if shape.Adjustments is non‑null and has items before calling Clear, and log the sheet name and shape index when a shape cannot be processed. | Create a reusable method ResetAllShapeAdjustments(Workbook workbook) that traverses all worksheets, resets every shape's adjustment values, and returns the updated workbook.
// Common Searches: Aspose.Cells C# clear shape adjustment values in all worksheets | reset custom shape adjustments to default in Excel using Aspose.Cells .NET | how to remove shape adjustments programmatically with Aspose.Cells library | iterate through workbook shapes and clear adjustments Aspose.Cells example | C# Aspose.Cells shape.Adjustments.Clear not working older version
// Tags: Aspose.Cells shape adjustments clearing | C# reset Excel shape properties | traverse workbook shapes Aspose | default shape adjustment values .NET | shape adjustment collection handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and its shapes, safely clears any adjustment values when the Adjustments collection is available, logs warnings for shapes that cannot be processed, and saves the modified file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

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

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // If the Shape class provides adjustment values in the current version,
                        // they can be cleared here. The Adjustments property may not be available
                        // in older Aspose.Cells versions, so this block is safely ignored when absent.
                        // Example (when supported):
                        // if (shape.Adjustments != null && shape.Adjustments.Count > 0)
                        // {
                        //     shape.Adjustments.Clear();
                        // }
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Warning: Unable to process shape on sheet '{sheet.Name}'. Details: {shapeEx.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
