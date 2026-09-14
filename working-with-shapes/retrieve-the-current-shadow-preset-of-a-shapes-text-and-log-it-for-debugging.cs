// Title: How to read and log a shape's text shadow preset in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, selects the first shape on the first worksheet, and prints its TextEffectFormat.ShadowPreset value to the console. | Create a reusable C# method that iterates over all shapes in a worksheet, safely accesses each shape's TextEffectFormat via dynamic typing, and logs the shadow preset for every shape, handling missing TextEffectFormat support. | Demonstrate how to catch a RuntimeBinderException when the Aspose.Cells version does not expose TextEffectFormat on a shape and output a clear fallback message.
// Common Searches: aspnet get text shadow preset from shape using Aspose.Cells | c# read shape TextEffectFormat shadow preset in Excel file | how to log shape text effect properties with Aspose.Cells .NET | dynamic access TextEffectFormat on shape Aspose.Cells version compatibility | retrieve shape shadow preset from worksheet with Aspose.Cells
// Tags: Aspose.Cells read shape TextEffectFormat | C# get shape text shadow preset | log shape text effect properties Aspose.Cells | dynamic TextEffectFormat access Aspose.Cells | handle RuntimeBinderException Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an .xlsx workbook, verifies that at least one shape exists, uses dynamic typing to read the shape's TextEffectFormat.ShadowPreset, writes the preset to the console, and gracefully handles RuntimeBinderException for older Aspose.Cells versions.
class Program
{
    static void Main()
    {
        try
        {
            const string filePath = "input.xlsx";

            // Verify the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Get the first worksheet (or specify another index/name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Access the first shape (adjust index or use name lookup as required)
            Shape shape = worksheet.Shapes[0];

            try
            {
                // Use dynamic to access TextEffectFormat at runtime (avoids compile‑time errors if the property is unavailable in the referenced version)
                dynamic dynShape = shape;
                var shadowPreset = dynShape.TextEffectFormat.ShadowPreset;

                // Output the retrieved shadow preset
                Console.WriteLine($"Text shadow preset: {shadowPreset}");
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("The loaded Aspose.Cells version does not support TextEffectFormat on Shape.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
