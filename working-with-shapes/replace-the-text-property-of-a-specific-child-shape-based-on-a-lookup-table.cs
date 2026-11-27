// Title: Replace specific child shape text in an Excel worksheet using Aspose.Cells for .NET with a case‑insensitive lookup dictionary
// Common Searches: Aspose.Cells C# replace shape text using a lookup dictionary | Update text of grouped shapes in an Excel file with Aspose.Cells | Iterate over worksheet shapes and change their Text property in .NET | How to change child shape text based on old‑new mapping in Aspose.Cells | Replace shape Text values in Excel workbook programmatically with C#
// Tags: shape text replacement Aspose.Cells C# | grouped shape text update Aspose.Cells | shape text lookup dictionary .NET | worksheet shapes iteration Aspose.Cells | excel shape content modification C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, defines a case‑insensitive dictionary mapping old shape text to new text, iterates through all worksheet shapes (including nested shapes inside group shapes accessed via reflection), replaces each shape's Text property when it matches a dictionary key, and saves the updated workbook.
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
            Worksheet sheet = workbook.Worksheets[0];

            // Define the lookup table: key = existing text, value = new text
            var lookupTable = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "OldText1", "NewText1" },
                { "OldText2", "NewText2" }
                // add more entries as needed
            };

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Process group shapes via reflection (covers API variations)
                    if (shape is GroupShape groupShape)
                    {
                        var innerShapes = GetGroupShapeItems(groupShape);
                        if (innerShapes != null)
                        {
                            foreach (Shape inner in innerShapes)
                            {
                                ReplaceShapeText(inner, lookupTable);
                            }
                            continue; // group shape processed, skip outer handling
                        }
                    }

                    // Process regular (or unsupported group) shapes
                    ReplaceShapeText(shape, lookupTable);
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Error processing shape: {exShape.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Retrieve inner shapes of a group shape using reflection (covers API variations)
    private static IEnumerable<Shape> GetGroupShapeItems(GroupShape groupShape)
    {
        try
        {
            if (groupShape == null) return null;

            // Attempt to get the "GroupShapeItems" property via reflection
            var prop = groupShape.GetType().GetProperty("GroupShapeItems");
            if (prop != null)
            {
                return prop.GetValue(groupShape) as IEnumerable<Shape>;
            }

            // Fallback: some versions expose a "Shapes" collection
            var shapesProp = groupShape.GetType().GetProperty("Shapes");
            if (shapesProp != null)
            {
                return shapesProp.GetValue(groupShape) as IEnumerable<Shape>;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to retrieve group shape items: {ex.Message}");
        }

        return null;
    }

    // Replace shape text based on the lookup table
    private static void ReplaceShapeText(Shape shape, Dictionary<string, string> lookup)
    {
        if (shape == null) return;

        string currentText = shape.Text;
        if (!string.IsNullOrEmpty(currentText) && lookup.TryGetValue(currentText, out string newText))
        {
            shape.Text = newText;
        }
    }
}
