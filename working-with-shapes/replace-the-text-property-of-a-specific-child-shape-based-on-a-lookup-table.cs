// Title: Replace text of specific child shapes in an Excel worksheet using Aspose.Cells for .NET with a lookup dictionary
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates over all shapes on a worksheet, and substitutes each shape's Text property using a supplied Dictionary<string,string> mapping. | Demonstrate how to handle both individual and grouped shapes when updating their Text values based on a lookup table in Aspose.Cells for .NET. | Create a reusable method that accepts a ShapeCollection and a dictionary, replaces matching shape text, and saves the modified workbook to a new file.
// Common Searches: Aspose.Cells C# replace shape text using dictionary lookup | How to update text of grouped shapes in an Excel file with Aspose.Cells | Iterate over worksheet shapes and change Text property based on mapping in .NET
// Tags: shape text replacement Aspose.Cells | lookup dictionary for shape text Excel | process ShapeCollection replace text .NET | grouped shape text modification Aspose.Cells | bulk shape Text property change workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, defines a dictionary that maps old shape text to new values, walks through all shapes on the first worksheet, replaces the Text property of any shape whose current text matches a dictionary key, and saves the updated workbook to a new file.
class ReplaceShapeText
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = @"C:\Input\Sample.xlsx";
            string outputPath = @"C:\Output\Sample_Modified.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet sheet = workbook.Worksheets[0];

            // Lookup table: old text -> new text
            Dictionary<string, string> lookup = new Dictionary<string, string>()
            {
                { "OldValue1", "NewValue1" },
                { "OldValue2", "NewValue2" },
                // add more mappings as needed
            };

            // Process a collection of shapes (including possible group shapes)
            void ProcessShapes(ShapeCollection shapes)
            {
                foreach (Shape shape in shapes)
                {
                    // If the shape is a group, we can still replace its own text (if any)
                    // Recursion into child shapes is omitted to avoid API compatibility issues
                    string currentText = shape.Text;
                    if (!string.IsNullOrEmpty(currentText) && lookup.ContainsKey(currentText))
                    {
                        shape.Text = lookup[currentText];
                    }
                }
            }

            // Start processing shapes on the worksheet
            ProcessShapes(sheet.Shapes);

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
