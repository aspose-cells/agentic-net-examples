// Title: Identify SmartArt shapes in an Excel worksheet by iterating the Shapes collection with Aspose.Cells for .NET (C#)
// Common Searches: C# code to list SmartArt shapes in an Excel file using Aspose.Cells | how to check if a worksheet shape is SmartArt with Aspose.Cells .NET | enumerate all shapes in a worksheet and find SmartArt objects Aspose.Cells | Aspose.Cells Shape.IsSmartArt example for .NET developers | retrieve SmartArt shape names and indexes from an XLSX using Aspose.Cells
// Tags: Aspose.Cells SmartArt detection | Shape.IsSmartArt API | worksheet shapes iteration Aspose.Cells | SmartArt objects extraction from worksheet | C# Aspose.Cells shape enumeration | extract SmartArt name index .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, accesses the first worksheet, iterates through its Shapes collection, uses the Shape.IsSmartArt property to spot SmartArt objects, prints each SmartArt shape's name and collection index, and finally saves the workbook unchanged.
class SmartArtIdentifier
{
    static void Main()
    {
        // Define input and output file paths (replace placeholders with actual paths)
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Check if the shape is a SmartArt object
                if (shape.IsSmartArt)
                {
                    // Get the shape's index within the collection
                    int shapeIndex = worksheet.Shapes.IndexOf(shape);
                    Console.WriteLine($"SmartArt found: Name = {shape.Name}, Index = {shapeIndex}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (no modifications made in this example)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
