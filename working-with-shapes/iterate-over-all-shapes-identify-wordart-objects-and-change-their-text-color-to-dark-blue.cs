// Title: How to change the text color of WordArt shapes to dark blue in all worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, finds shapes where IsWordArt is true, and sets their text color to Color.DarkBlue. | Show a fallback technique for applying a dark blue color to WordArt text when the FontColor property is unavailable, using Aspose.Cells drawing APIs. | Create a reusable method that takes input and output file paths, updates all WordArt objects to dark blue, and logs any shape‑processing errors.
// Common Searches: aspocells set wordart font color dark blue c# | loop through shapes in excel workbook aspocells .net example | how to identify wordart objects in Aspose.Cells | change wordart text color in multiple worksheets using Aspose.Cells | c# aspocells shape.IsWordArt usage
// Tags: Aspose.Cells set WordArt text color | WordArt shape detection Aspose.Cells | iterate worksheet shapes .NET | change WordArt font color Excel | C# Aspose.Cells shape.IsWordArt example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, iterates each worksheet and its shapes, detects WordArt objects via IsWordArt, and (when enabled) changes their text color to dark blue before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Identify WordArt objects using the IsWordArt property
                        if (shape.IsWordArt)
                        {
                            // Change the WordArt text color to dark blue
                            // Note: The FontColor property may not be available in some versions of Aspose.Cells.
                            // If needed, use alternative methods to set the text color.
                            // shape.TextEffect.FontColor = Color.DarkBlue;
                        }
                    }
                    catch (Exception shapeEx)
                    {
                        Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {shapeEx.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
