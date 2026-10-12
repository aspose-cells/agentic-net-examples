// Title: How to filter and process only SmartArt shapes in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops over the worksheet.Shapes collection and runs custom logic exclusively for shapes where IsSmartArt is true, using the Aspose.Cells API. | Provide an example that prints the names of SmartArt objects in a workbook while skipping all other shape types with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# list SmartArt shapes in a worksheet | filter SmartArt objects in an Excel file using Aspose.Cells .NET | C# iterate worksheet shapes and filter by IsSmartArt property | how to identify SmartArt shapes with Aspose.Cells API | example to process only SmartArt shapes in an Excel workbook using Aspose.Cells
// Tags: filter SmartArt shapes Aspose.Cells | enumerate worksheet shapes C# | check IsSmartArt property Aspose.Cells | process only SmartArt objects Excel .NET | Aspose.Cells shape iteration example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates over all shapes on the first worksheet, selects only those where IsSmartArt is true, outputs their names, and saves the workbook unchanged.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook (lifecycle rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    // Example processing: output the shape name
                    Console.WriteLine($"SmartArt Shape: {shape.Name}");
                    // Additional SmartArt-specific logic can be placed here
                }
            }

            // Save the workbook (lifecycle rule) – unchanged if no modifications were made
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
