// Title: Ungroup a GroupShape, reposition each child shape, and save the workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells ungroup GroupShape and adjust child shape positions | How to move individual shapes after ungrouping a shape group in Aspose.Cells | Saving an Excel file after modifying shape coordinates with Aspose.Cells .NET | Aspose.Cells example to shift shapes 10 points after ungrouping | Retrieve and ungroup the first GroupShape in a worksheet using Aspose.Cells C#
// Tags: Aspose.Cells ungroup GroupShape C# | modify shape Left Top properties Aspose.Cells | shift Excel shapes after ungroup Aspose.Cells | save workbook with updated shapes Aspose.Cells | iterate worksheet shapes Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, finds the first GroupShape on the first worksheet, calls Ungroup() to separate it into individual shapes, moves each resulting shape 10 points to the right and down by updating its Left and Top properties, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Locate the first shape that is a group
            GroupShape groupShape = null;
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape is GroupShape gs)
                {
                    groupShape = gs;
                    break;
                }
            }

            if (groupShape != null)
            {
                try
                {
                    // Ungroup the shape; the shapes become independent on the worksheet
                    groupShape.Ungroup();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to ungroup shape: {ex.Message}");
                    return;
                }

                // After ungrouping, shift each resulting shape 10 points right and down
                foreach (Shape s in sheet.Shapes)
                {
                    // Skip any remaining group shapes (if any)
                    if (s is GroupShape) continue;

                    s.Left += 10; // move horizontally
                    s.Top  += 10; // move vertically
                }
            }
            else
            {
                Console.WriteLine("No group shape found in the worksheet.");
            }

            // Save the workbook with the updated shapes
            try
            {
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
