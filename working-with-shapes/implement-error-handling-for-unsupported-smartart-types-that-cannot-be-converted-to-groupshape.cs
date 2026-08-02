// Title: Handle Unsupported SmartArt Types When Converting to GroupShape with Aspose.Cells for .NET
// Description: Demonstrates how to load an Excel file, iterate through worksheets, detect SmartArt shapes, safely convert them to GroupShape using GetResultOfSmartArt, log failures or null results, and save the workbook with UpdateSmartArt enabled.
// Keywords: Aspose.Cells SmartArt conversion | GetResultOfSmartArt C# | unsupported SmartArt handling | GroupShape conversion error | Excel shape processing .NET
// Common Searches: Aspose.Cells how to skip SmartArt that cannot be converted | GetResultOfSmartArt returns null what to do | C# catch exception for unsupported SmartArt in Aspose.Cells | convert SmartArt to GroupShape with error handling | update SmartArt while saving workbook Aspose.Cells
// Developer Intent: Add robust checks and exception handling so that SmartArt shapes that cannot be turned into GroupShape do not crash the application.
// Use Cases: Log the name and error of each SmartArt shape that fails conversion and continue processing the rest. | Skip unsupported SmartArt types while applying modifications only to convertible shapes. | Replace non‑convertible SmartArt with a placeholder image or comment to keep the layout intact.
// AI Prompts: Generate C# code that verifies shape.IsSmartArt, attempts GetResultOfSmartArt inside a try‑catch, catches NotSupportedException or generic Exception, logs shape.Name and the error, and proceeds to the next shape. | Create a helper method IsSmartArtConvertible(Shape shape) that returns a boolean after attempting conversion and null checks, then integrate it into a loop that processes only convertible SmartArt shapes.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to load an Excel file, iterate through worksheets, detect SmartArt shapes, safely convert them to GroupShape using GetResultOfSmartArt, log failures or null results, and save the workbook with UpdateSmartArt enabled.
class SmartArtConversionDemo
{
    static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Iterate through all worksheets and their shapes
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    try
                    {
                        // Convert SmartArt to a GroupShape
                        GroupShape groupShape = shape.GetResultOfSmartArt();

                        if (groupShape != null)
                        {
                            // Example modification: shift the group slightly to the right
                            groupShape.Left += 10;
                        }
                        else
                        {
                            Console.WriteLine($"SmartArt shape '{shape.Name}' could not be converted (null result).");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting SmartArt shape '{shape.Name}': {ex.Message}");
                    }
                }
            }
        }

        // Save the workbook with SmartArt update enabled
        try
        {
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                UpdateSmartArt = true
            };
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
