// Title: Apply a 12‑point orange glow to the first shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with Aspose.Cells, locate the first worksheet shape, and set its EffectFormat.Glow.Size to 12 points and Glow.Color to orange, using dynamic binding for version compatibility. | Write C# code that verifies a shape exists in a workbook, adds a 12‑point orange glow via Shape.EffectFormat, and saves the file while gracefully handling missing EffectFormat support.
// Common Searches: Aspose.Cells C# set shape glow size to 12 points | how to change shape glow color to orange in Excel using Aspose.Cells | apply glow effect to Excel shape with Aspose.Cells .NET dynamic binding | C# code to add orange glow to first shape in workbook using Aspose.Cells | EffectFormat not supported fallback Aspose.Cells shape glow
// Tags: Aspose.Cells shape glow customization | set glow size points Aspose.Cells | change shape glow color Aspose.Cells | dynamic EffectFormat fallback C# | C# modify Excel shape appearance

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, checks for at least one shape on the first worksheet, and uses dynamic binding to set the shape's EffectFormat.Glow.Size to 12 points and Glow.Color to orange. It catches a RuntimeBinderException if the EffectFormat property is unavailable, then saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape (adjust index if you target a specific shape)
            Shape shape = worksheet.Shapes[0];

            // Attempt to apply glow effect using dynamic to handle versions without EffectFormat
            try
            {
                dynamic dynShape = shape;
                dynShape.EffectFormat.Glow.Size = 12f;          // Set glow size (points)
                dynShape.EffectFormat.Glow.Color = Color.Orange; // Set glow color
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // EffectFormat not supported in this version; skip effect
                Console.WriteLine("Glow effect not supported in the current Aspose.Cells version.");
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
}
