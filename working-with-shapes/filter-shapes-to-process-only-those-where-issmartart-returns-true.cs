// Title: Filtering worksheet shapes to process only SmartArt objects with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate worksheet shapes and select only SmartArt | how to check IsSmartArt property for shapes in Aspose.Cells | example code to log SmartArt shape name and type using Aspose.Cells | filter SmartArt shapes from ShapeCollection in Excel file with Aspose.Cells | process only SmartArt objects in a workbook using Aspose.Cells .NET
// Tags: SmartArt shape filtering Aspose.Cells | Iterate worksheet ShapeCollection C# | IsSmartArt property usage .NET | Log SmartArt name and type Aspose.Cells | Aspose.Cells workbook shape processing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtFilter
{
    // The sample loads an Excel workbook with Aspose.Cells, accesses the first worksheet, iterates through its ShapeCollection, and processes only shapes where IsSmartArt is true. For each SmartArt shape it logs the shape's Name and Type, then saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve all shapes on the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Process only SmartArt shapes
                foreach (Shape shape in shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        try
                        {
                            // Since Aspose.Cells does not expose a direct SmartArt object,
                            // we can log basic information about the SmartArt shape.
                            Console.WriteLine($"SmartArt Shape Name: {shape.Name}");
                            Console.WriteLine($"Shape Type: {shape.Type}");
                            // Additional SmartArt processing can be added here if supported in future versions.
                        }
                        catch (Exception ex)
                        {
                            // Handle any issues retrieving SmartArt details
                            Console.WriteLine($"Error processing SmartArt shape '{shape.Name}': {ex.Message}");
                        }
                    }
                }

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook after processing
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
