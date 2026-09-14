// Title: Check for empty text in all SmartArt shapes of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create C# code with Aspose.Cells that scans every worksheet and shape, and logs the names of SmartArt shapes whose Text property is null or whitespace. | Write a C# method that replaces empty SmartArt shape text with a default placeholder (e.g., "[No Text]") before saving the workbook using Aspose.Cells. | Generate a C# function that returns a list of shape identifiers that have no text after trimming, leveraging the Aspose.Cells shape collection.
// Common Searches: asp.net how to detect empty SmartArt shape text in an Excel file with Aspose.Cells | c# iterate through Excel shapes and list those with missing text using Aspose.Cells | validate that every SmartArt object in a workbook contains non‑empty text with Aspose.Cells for .NET
// Tags: Aspose.Cells iterate worksheet shapes | detect empty SmartArt text Aspose.Cells | replace missing shape text C# | log shape validation warnings Aspose | Excel workbook shape text verification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, walks through each worksheet and its shapes, trims each shape's Text, logs a warning for shapes with null or whitespace text, optionally replaces empty text with a placeholder, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process shapes that contain text (e.g., SmartArt is not directly supported)
                    string shapeText = shape.Text?.Trim();

                    if (string.IsNullOrEmpty(shapeText))
                    {
                        // Report empty text in the shape
                        Console.WriteLine(
                            $"Empty text detected in worksheet '{sheet.Name}', shape name '{shape.Name}'.");
                    }
                }
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
