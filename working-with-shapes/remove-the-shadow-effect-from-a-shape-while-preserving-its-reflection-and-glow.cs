// Title: How to disable a shape’s shadow while keeping its reflection and glow in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide shape shadow but retain reflection effect | remove shadow from Excel shape without affecting glow using Aspose.Cells | disable shape shadow programmatically in a .NET workbook while preserving other effects | dynamic access to Shadow property in Aspose.Cells for different library versions | set shape shadow visibility false in Aspose.Cells C# example
// Tags: shape shadow visibility Aspose.Cells | shape reflection preservation Aspose.Cells | shape glow retention Aspose.Cells | dynamic shape property access Aspose.Cells | excel shape effect manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Microsoft.CSharp.RuntimeBinder;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, obtains the first shape on the first worksheet, uses a dynamic cast to safely access the shape's Shadow property across Aspose.Cells versions, sets IsVisible to false to turn off the shadow while leaving reflection and glow unchanged, and then saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one shape
                if (sheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet.");
                    return;
                }

                // Retrieve the first shape (or use sheet.Shapes["ShapeName"] if known)
                Shape shape = sheet.Shapes[0];

                // Use dynamic to access the Shadow property safely across different Aspose.Cells versions
                dynamic dynShape = shape;
                try
                {
                    // Check if the Shadow property exists and is not null
                    if (dynShape.Shadow != null)
                    {
                        // Disable the shadow while keeping other effects intact
                        dynShape.Shadow.IsVisible = false;
                    }
                }
                catch (RuntimeBinderException)
                {
                    // The Shadow property is not available in this version of Aspose.Cells
                    Console.WriteLine("Shadow property is not supported in the current Aspose.Cells version.");
                }

                // Save the workbook with the updated shape
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // General exception handling for unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
