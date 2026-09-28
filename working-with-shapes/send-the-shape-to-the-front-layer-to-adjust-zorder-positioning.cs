// Title: How to bring a specific shape to the front layer in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, locate a shape by its name (or fall back to the first shape), set its ZOrderPosition to the highest index, and save the file with Aspose.Cells in C#. | Programmatically adjust the Z‑order of a worksheet shape to the topmost layer using the Shape.ZOrderPosition property in Aspose.Cells for .NET. | Retrieve a shape from a worksheet, change its layering order to bring it forward, and write the updated workbook to a new file with the Aspose.Cells C# API.
// Common Searches: aspnet c# set shape ZOrderPosition to bring shape to front in Excel file | Aspose.Cells move shape to top layer programmatically | how to change shape layering order in an Excel worksheet using Aspose.Cells | retrieve shape by name and adjust Z order with Aspose.Cells .NET
// Tags: Aspose.Cells Shape.ZOrderPosition | C# set Excel shape Z-order | Aspose.Cells retrieve worksheet shape | Excel shape layering with Aspose.Cells | Aspose.Cells topmost shape positioning

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // This example loads an existing workbook, obtains a shape named "MyShape" (or the first shape if not found), sets its ZOrderPosition to the highest index to bring it to the front, and saves the modified workbook to a new file.
class ShapeZOrderExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the shape by name; fallback to first shape if not found
            Shape? shape = null;
            try
            {
                shape = worksheet.Shapes["MyShape"];
            }
            catch
            {
                // Ignored – will try alternative retrieval below
            }

            if (shape == null && worksheet.Shapes.Count > 0)
            {
                shape = worksheet.Shapes[0];
            }

            if (shape == null)
            {
                Console.WriteLine("No shapes available in the worksheet.");
                return;
            }

            // Bring the shape to the front by setting its Z‑order to the highest position
            shape.ZOrderPosition = worksheet.Shapes.Count - 1;

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
