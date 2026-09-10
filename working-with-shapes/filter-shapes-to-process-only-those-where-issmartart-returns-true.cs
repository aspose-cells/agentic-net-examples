// Title: How to filter and work with only SmartArt shapes in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates over worksheet.Shapes, and runs custom logic only when shape.IsSmartArt is true. | Provide a C# example that enumerates SmartArt shapes, prints each shape's Name and Type, and saves the workbook after processing.
// Common Searches: C# Aspose.Cells iterate worksheet shapes and check IsSmartArt property | How to list SmartArt objects in an Excel file using Aspose.Cells for .NET | Filter SmartArt shapes in an Aspose.Cells workbook before saving | Retrieve name and type of SmartArt shapes with Aspose.Cells C# example
// Tags: Aspose.Cells filter SmartArt shapes C# | Iterate worksheet shapes IsSmartArt Aspose.Cells | List SmartArt shape names and types Aspose.Cells | Process only SmartArt objects Excel .NET | SmartArt shape handling Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel file with Aspose.Cells, accesses the first worksheet, loops through all shapes, and processes only those where the IsSmartArt property is true, outputting each SmartArt shape's name and type before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    Console.WriteLine($"SmartArt Shape: Name = {shape.Name}, Type = {shape.Type}");
                    // Additional SmartArt-specific logic can be placed here
                }
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
