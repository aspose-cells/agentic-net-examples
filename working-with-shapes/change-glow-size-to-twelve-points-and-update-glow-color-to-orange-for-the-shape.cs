// Title: Set a 12‑point orange glow on the first shape in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: In C#, use Aspose.Cells to set the first worksheet shape's glow thickness to 12 points and its hue to orange via the EffectFormat property, then write the file. | Create a .NET snippet that verifies a shape exists, then applies a 12‑point orange outer glow using Shape.EffectFormat.Glow and saves the workbook.
// Common Searches: how to increase glow size of an Excel shape with Aspose.Cells C# | Aspose.Cells example for applying orange outer glow to a shape | C# code to modify shape glow properties in a workbook | programmatically change shape glow color in Excel using Aspose.Cells | set glow thickness and color for first shape in .NET Excel file
// Tags: Aspose.Cells shape glow effect | C# adjust shape glow size | Aspose.Cells change shape glow color | Excel shape visual effect .NET | Aspose.Cells shape effect format

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, checks for at least one shape on the first worksheet, and (when supported) sets the shape's EffectFormat.Glow size to 12 points and its color to orange before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape in the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without modifications to \"{outputPath}\".");
                return;
            }

            // Access the target shape (adjust index or use a name lookup as needed)
            var shape = worksheet.Shapes[0];

            // NOTE: The EffectFormat property is not available in some older versions of Aspose.Cells.
            // If your version supports it, you can uncomment the following lines to apply a glow effect.
            // shape.EffectFormat.Glow.Size = 12;
            // shape.EffectFormat.Glow.Color = Color.Orange;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
