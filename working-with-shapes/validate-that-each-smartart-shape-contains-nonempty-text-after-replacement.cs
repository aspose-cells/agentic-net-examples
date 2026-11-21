// Title: How to validate that every SmartArt shape contains non‑empty text in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate worksheet shapes and find SmartArt items with blank text | detect empty or whitespace SmartArt shape text in an Excel file with Aspose.Cells | list all shapes that have null or empty Text property in a workbook using Aspose.Cells for .NET
// Tags: Aspose.Cells iterate worksheet shapes | SmartArt shape text validation C# | detect empty shape text Aspose.Cells | list Excel shapes with blank text .NET | Excel workbook shape text check Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and its collection of shapes, reads the Shape.Text property, flags any shape whose text is null, empty, or only whitespace, and prints a report of the worksheet and shape names that contain invalid (empty) text.
class SmartArtValidator
{
    static void Main()
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // List to hold information about shapes with empty text
        List<string> invalidShapes = new List<string>();

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Retrieve the text of the shape; Shape.Text may be null
                    string shapeText = shape.Text ?? string.Empty;

                    // Validate that the text is not null, empty, or whitespace
                    if (string.IsNullOrWhiteSpace(shapeText))
                    {
                        // Record the worksheet name and shape name for reporting
                        invalidShapes.Add($"Worksheet: '{sheet.Name}', Shape: '{shape.Name}'");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during validation: {ex.Message}");
            return;
        }

        // Output validation results
        if (invalidShapes.Count == 0)
        {
            Console.WriteLine("All shapes contain non‑empty text.");
        }
        else
        {
            Console.WriteLine("The following shapes have empty or whitespace text:");
            foreach (string info in invalidShapes)
            {
                Console.WriteLine(info);
            }
        }
    }
}
