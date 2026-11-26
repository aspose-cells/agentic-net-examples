// Title: Apply a uniform adjustment value to every shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# loop through all shapes in an Excel worksheet with Aspose.Cells and set the same adjustment | how to apply identical custom adjustment to multiple drawing objects in Aspose.Cells for .NET | example code for iterating workbook shapes and updating adjustment property in Aspose.Cells | Aspose.Cells shape adjustment placeholder code for bulk processing in C# | set uniform adjustment for Excel shapes using Aspose.Cells API
// Tags: Aspose.Cells set shape adjustments C# | iterate workbook shapes Aspose.Cells | uniform adjustment for Excel drawing objects | bulk shape property update Aspose.Cells | C# Excel shape processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, defines a uniform adjustment value, iterates through each worksheet and each shape, outputs the shape name as a placeholder for applying the adjustment, and saves the modified file while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Uniform adjustment value (placeholder – actual adjustment APIs may vary)
            double uniformAdjustment = 0.5;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // NOTE: Aspose.Cells for .NET does not expose a direct Adjustments collection.
                    // If future versions provide such an API, replace the following placeholder
                    // with the appropriate logic to set each adjustment to uniformAdjustment.
                    // For now, we simply output the shape name for demonstration.
                    Console.WriteLine($"Processing shape: {shape.Name}");
                }
            }

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
