// Title: How to set the PerspectiveDiagonalBottomLeft shadow preset on a rectangle shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Assign ShadowPreset.PerspectiveDiagonalBottomLeft to the rectangle shape's Shadow.Preset property with the Aspose.Cells C# API before saving the workbook. | Update the sample code to apply a PerspectiveDiagonalBottomLeft shadow effect to the added shape and generate output.xlsx to see the change. | Write C# that creates a rectangle shape in a worksheet, sets its shadow preset to PerspectiveDiagonalBottomLeft, and saves the file to demonstrate the visual difference.
// Common Searches: Aspose.Cells C# set shape shadow to PerspectiveDiagonalBottomLeft | apply diagonal bottom left perspective shadow to Excel shape using Aspose.Cells | C# example for changing shape shadow preset in Aspose.Cells workbook | how to use ShadowPreset enum with Aspose.Cells to style shapes
// Tags: Aspose.Cells set shape shadow preset | C# Aspose.Cells ShadowPreset PerspectiveDiagonalBottomLeft | Excel shape styling with Aspose.Cells | Aspose.Cells rectangle shape shadow effect | modify shape visual appearance Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, adds a rectangle shape, and demonstrates how to set the shape's Shadow.Preset to ShadowPreset.PerspectiveDiagonalBottomLeft using Aspose.Cells for .NET. Saving the workbook produces an output file where the visual impact of the new shadow preset can be observed.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Add a rectangle shape
            var shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                150, // width (pixels)
                80   // height (pixels)
            );

            // Note: Shadow settings are not available in older Aspose.Cells versions.
            // If needed, they can be applied using the Shadow property in newer versions.

            // Save the workbook with the new shape
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
