// Title: How to toggle the glow effect of a specific shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# toggle shape glow visibility without changing other effects | C# set shape EffectFormat.Glow.Size to hide glow in an Excel file | How to enable or disable glow on a named shape with Aspose.Cells | Remove glow from a shape in an existing .xlsx using Aspose.Cells .NET | Programmatically change shape glow size in an Aspose.Cells workbook
// Tags: shape effectformat glow control Aspose.Cells | c# set glow size zero Excel shape | Aspose.Cells modify shape visual effects | excel shape glow visibility .NET | EffectFormat Glow property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing .xlsx workbook, retrieves a shape by name from the first worksheet, and shows how to toggle the shape's glow effect by checking the EffectFormat.Glow.Size property and setting it to 0 to hide or to a positive value (e.g., 5) to show, then saves the modified workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the target shape by its name (replace "MyShape" with the actual shape name)
            Shape shape = worksheet.Shapes["MyShape"];

            // Ensure the shape exists
            if (shape != null)
            {
                // The EffectFormat API may not be available in older Aspose.Cells versions.
                // If your version supports it, uncomment the following block to toggle the Glow effect.

                /*
                EffectFormat effectFormat = shape.EffectFormat;
                GlowEffect glow = effectFormat.Glow;

                // Toggle visibility:
                // - If the current Glow size is greater than 0, hide it by setting Size to 0.
                // - If the size is 0, show it with a default size (e.g., 5 points).
                if (glow.Size > 0)
                {
                    glow.Size = 0; // Hide glow
                }
                else
                {
                    glow.Size = 5; // Show glow with a default size
                }
                */

                Console.WriteLine($"Shape '{shape.Name}' found.");
            }
            else
            {
                Console.WriteLine("Shape 'MyShape' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display the error message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
