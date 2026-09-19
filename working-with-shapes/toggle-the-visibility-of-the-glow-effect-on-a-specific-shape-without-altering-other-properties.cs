// Title: Toggle a shape's glow effect on or off in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that flips the IsEnabled property of a shape's Glow effect using Aspose.Cells while leaving all other shape formatting unchanged. | Show how to use dynamic binding to access EffectFormat and safely toggle a shape's glow visibility when the API may be missing in certain Aspose.Cells versions. | Create a complete example that loads an .xlsx file, retrieves the first shape, toggles its glow effect, and saves the modified workbook.
// Common Searches: how to enable or disable glow effect on a shape with Aspose.Cells C# | Aspose.Cells toggle shape glow without affecting other formatting | C# dynamic binding for shape EffectFormat in Aspose.Cells | compatible way to change shape glow visibility across Aspose.Cells versions | sample code to turn off shape glow in an Excel file using Aspose.Cells
// Tags: shape glow toggle Aspose.Cells | Aspose.Cells dynamic EffectFormat access | C# toggle shape glow visibility | Excel shape glow enable disable | version‑agnostic shape effect handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, selects the first shape on the first worksheet, uses dynamic binding to flip the Glow.IsEnabled flag (turning the glow effect on or off) while preserving all other shape properties, handles cases where the EffectFormat API is unavailable, and saves the updated file.
class ToggleShapeGlow
{
    static void Main()
    {
        try
        {
            // Verify input file exists
            string inputPath = @"C:\Input\Sample.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = sheet.Shapes[0];

            // Attempt to toggle the glow effect using dynamic binding.
            // This avoids compile‑time errors if the EffectFormat/Glow API is unavailable
            // in the referenced Aspose.Cells version.
            try
            {
                dynamic dynShape = shape;
                var effectFormat = dynShape.EffectFormat;   // May throw RuntimeBinderException
                var glow = effectFormat.Glow;               // May throw RuntimeBinderException
                glow.IsEnabled = !glow.IsEnabled;           // Toggle visibility
                Console.WriteLine("Glow effect toggled successfully.");
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("Glow effect is not supported in the current Aspose.Cells version.");
            }

            // Verify output directory exists
            string outputPath = @"C:\Output\Sample_ToggledGlow.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
